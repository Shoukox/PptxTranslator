using System.Collections.Concurrent;
using System.ComponentModel;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.Mvc;

namespace PptxTranslator.Web.Features.Translation;

internal static partial class TranslateFeature
{
    private static readonly ConcurrentDictionary<string, TranslationJob> Jobs = new();
    private static readonly TimeSpan JobRetention = TimeSpan.FromHours(1);

    public static IEndpointRouteBuilder MapTranslateFeature(this IEndpointRouteBuilder endpoints)
    {
        endpoints.MapPost("/translate/jobs", HandleCreateJobAsync);
        endpoints.MapGet("/translate/jobs/{jobId}", HandleGetJobAsync);
        endpoints.MapGet("/translate/jobs/{jobId}/download", HandleDownloadJobResultAsync);
        return endpoints;
    }

    [RequestSizeLimit(1024 * 1024 * 1024)] // 1 GB
    [RequestFormLimits(MultipartBodyLengthLimit = 1024 * 1024 * 1024)] // 1 GB
    private static async Task<IResult> HandleCreateJobAsync(
        HttpRequest request,
        IWebHostEnvironment environment,
        ILogger<Program> logger,
        CancellationToken cancellationToken)
    {
        CleanupExpiredJobs(logger);

        if (!request.HasFormContentType)
        {
            return Results.BadRequest("Expected multipart/form-data.");
        }

        var form = await request.ReadFormAsync(cancellationToken);
        var file = form.Files.GetFile("pptxFile");

        if (file is null || file.Length == 0)
        {
            return Results.BadRequest("Please choose a .pptx file to translate.");
        }

        if (!IsPptxFile(file.FileName))
        {
            return Results.BadRequest("Only .pptx files are supported.");
        }

        var sourceLanguage = NormalizeLanguage(form["sourceLanguage"].ToString(), "en");
        var targetLanguage = NormalizeLanguage(form["targetLanguage"].ToString(), "ru");
        var workingDirectory = CreateWorkingDirectory();

        TranslationRequestFiles requestFiles;
        try
        {
            requestFiles = await SaveRequestFilesAsync(file, targetLanguage, workingDirectory, cancellationToken);
        }
        catch
        {
            DeleteWorkingDirectory(workingDirectory, logger);
            throw;
        }

        var scriptPath = ResolveScriptPath(environment.ContentRootPath);
        var createdAtUtc = DateTimeOffset.UtcNow;
        var job = new TranslationJob(
            Id: Guid.NewGuid().ToString("n"),
            Status: TranslationJobStatus.Queued,
            SourceLanguage: sourceLanguage,
            TargetLanguage: targetLanguage,
            WorkingDirectory: workingDirectory,
            InputPath: requestFiles.InputPath,
            OutputPath: requestFiles.OutputPath,
            OutputFileName: requestFiles.OutputFileName,
            CreatedAtUtc: createdAtUtc,
            UpdatedAtUtc: createdAtUtc,
            ExpiresAtUtc: createdAtUtc.Add(JobRetention),
            PartialFailureCount: 0,
            ErrorMessage: null,
            PythonCommand: null,
            StandardOutput: null,
            StandardError: null);

        Jobs[job.Id] = job;

        _ = Task.Run(async () =>
        {
            UpdateJob(job.Id, current => current with
            {
                Status = TranslationJobStatus.Running,
                UpdatedAtUtc = DateTimeOffset.UtcNow
            });

            try
            {
                var translationResult = await RunTranslationAsync(
                    job.InputPath,
                    job.OutputPath,
                    scriptPath,
                    job.SourceLanguage,
                    job.TargetLanguage,
                    job.WorkingDirectory,
                    CancellationToken.None);

                if (!translationResult.Succeeded)
                {
                    logger.LogError(
                        "Translation failed. JobId: {JobId}, ExitCode: {ExitCode}, Python: {PythonCommand}, StdOut: {StdOut}, StdErr: {StdErr}",
                        job.Id,
                        translationResult.ExitCode,
                        translationResult.PythonCommand,
                        translationResult.StandardOutput,
                        translationResult.StandardError);

                    UpdateJob(job.Id, current => current with
                    {
                        Status = TranslationJobStatus.Failed,
                        UpdatedAtUtc = DateTimeOffset.UtcNow,
                        ErrorMessage = "Something went wrong...",
                        PartialFailureCount = translationResult.PartialFailureCount,
                        PythonCommand = translationResult.PythonCommand,
                        StandardOutput = translationResult.StandardOutput,
                        StandardError = translationResult.StandardError
                    });

                    return;
                }

                logger.LogInformation(
                    "Translation succeeded. JobId: {JobId}, Python: {PythonCommand}, StdOut: {StdOut}, StdErr: {StdErr}",
                    job.Id,
                    translationResult.PythonCommand,
                    translationResult.StandardOutput,
                    translationResult.StandardError);

                UpdateJob(job.Id, current => current with
                {
                    Status = TranslationJobStatus.Completed,
                    UpdatedAtUtc = DateTimeOffset.UtcNow,
                    PartialFailureCount = translationResult.PartialFailureCount,
                    PythonCommand = translationResult.PythonCommand,
                    StandardOutput = translationResult.StandardOutput,
                    StandardError = translationResult.StandardError
                });
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Translation job {JobId} crashed.", job.Id);
                UpdateJob(job.Id, current => current with
                {
                    Status = TranslationJobStatus.Failed,
                    UpdatedAtUtc = DateTimeOffset.UtcNow,
                    ErrorMessage = "Something went wrong..."
                });
            }
        });

        return Results.Accepted($"/translate/jobs/{job.Id}", new TranslationJobResponse(
            job.Id,
            job.Status.ToApiValue(),
            null,
            null,
            null,
            null,
            job.CreatedAtUtc,
            job.UpdatedAtUtc,
            job.ExpiresAtUtc));
    }

    private static IResult HandleGetJobAsync(string jobId, ILogger<Program> logger)
    {
        CleanupExpiredJobs(logger);

        if (!Jobs.TryGetValue(jobId, out var job))
        {
            return Results.NotFound();
        }

        return Results.Ok(ToResponse(job));
    }

    private static IResult HandleDownloadJobResultAsync(HttpContext httpContext, string jobId, ILogger<Program> logger)
    {
        CleanupExpiredJobs(logger);

        if (!Jobs.TryGetValue(jobId, out var job))
        {
            return Results.NotFound();
        }

        if (job.Status != TranslationJobStatus.Completed)
        {
            return Results.BadRequest("Translation is not ready yet.");
        }

        if (!File.Exists(job.OutputPath))
        {
            logger.LogWarning("Completed translation job {JobId} is missing output file {OutputPath}.", jobId, job.OutputPath);
            return Results.Problem(
                title: "Translation result is unavailable",
                detail: "The translated file is no longer available.",
                statusCode: StatusCodes.Status410Gone);
        }

        if (job.PartialFailureCount > 0)
        {
            httpContext.Response.Headers.Append(
                "X-Translation-Notice",
                "Translation completed. Some text was kept in the original language.");
        }

        return Results.File(
            job.OutputPath,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            job.OutputFileName,
            enableRangeProcessing: false);
    }

    private static TranslationJobResponse ToResponse(TranslationJob job) => new(
        job.Id,
        job.Status.ToApiValue(),
        job.Status switch
        {
            TranslationJobStatus.Queued => "Queued for translation.",
            TranslationJobStatus.Running => "Translating presentation...",
            TranslationJobStatus.Completed => "Translation complete.",
            TranslationJobStatus.Failed => job.ErrorMessage ?? "Translation failed.",
            _ => "Unknown status."
        },
        job.Status == TranslationJobStatus.Completed ? $"/translate/jobs/{job.Id}/download" : null,
        job.OutputFileName,
        job.PartialFailureCount,
        job.CreatedAtUtc,
        job.UpdatedAtUtc,
        job.ExpiresAtUtc);

    private static void UpdateJob(string jobId, Func<TranslationJob, TranslationJob> update)
    {
        while (Jobs.TryGetValue(jobId, out var current))
        {
            var next = update(current);
            if (Jobs.TryUpdate(jobId, next, current))
            {
                return;
            }
        }
    }

    private static void CleanupExpiredJobs(ILogger logger)
    {
        var now = DateTimeOffset.UtcNow;

        foreach (var entry in Jobs)
        {
            if (entry.Value.ExpiresAtUtc > now)
            {
                continue;
            }

            if (Jobs.TryRemove(entry.Key, out var removed))
            {
                DeleteWorkingDirectory(removed.WorkingDirectory, logger);
            }
        }
    }

    private static bool IsPptxFile(string fileName) =>
        string.Equals(Path.GetExtension(fileName), ".pptx", StringComparison.OrdinalIgnoreCase);

    private static string CreateWorkingDirectory()
    {
        var workingDirectory = Path.Combine(Path.GetTempPath(), "pptx-translator", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workingDirectory);
        return workingDirectory;
    }

    private static async Task<TranslationRequestFiles> SaveRequestFilesAsync(
        IFormFile file,
        string targetLanguage,
        string workingDirectory,
        CancellationToken cancellationToken)
    {
        var safeInputName = Path.GetFileName(file.FileName);
        var inputPath = Path.Combine(workingDirectory, safeInputName);

        await using (var inputStream = File.Create(inputPath))
        {
            await file.CopyToAsync(inputStream, cancellationToken);
        }

        var outputFileName = $"{Path.GetFileNameWithoutExtension(safeInputName)}_{targetLanguage}.pptx";
        var outputPath = Path.Combine(workingDirectory, outputFileName);

        return new TranslationRequestFiles(inputPath, outputPath, outputFileName);
    }

    private static void DeleteWorkingDirectory(string workingDirectory, ILogger logger)
    {
        try
        {
            if (Directory.Exists(workingDirectory))
            {
                Directory.Delete(workingDirectory, recursive: true);
            }
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "Failed to delete temporary translation directory {Directory}", workingDirectory);
        }
    }

    private static string NormalizeLanguage(string? value, string fallback)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return fallback;
        }

        return value.Trim() switch
        {
            "English" or "english" => "en",
            "Russian" or "russian" => "ru",
            "German" or "german" => "de",
            "French" or "french" => "fr",
            "Spanish" or "spanish" => "es",
            "Italian" or "italian" => "it",
            "Portuguese" or "portuguese" => "pt",
            "Polish" or "polish" => "pl",
            "Ukrainian" or "ukrainian" => "uk",
            "Turkish" or "turkish" => "tr",
            "Dutch" or "dutch" => "nl",
            "Czech" or "czech" => "cs",
            "Romanian" or "romanian" => "ro",
            "Japanese" or "japanese" => "ja",
            "Korean" or "korean" => "ko",
            "Chinese (Simplified)" or "chinese (simplified)" => "zh-CN",
            _ => value.Trim()
        };
    }

    private static string ResolveScriptPath(string contentRootPath)
    {
        var bundledScript = Path.Combine(contentRootPath, "Scripts", "pptx_translate_ru.py");
        const string hostScript = @"C:\Users\shach\Desktop\translator\pptx_translate_ru.py";

        if (File.Exists(hostScript))
        {
            return hostScript;
        }

        if (File.Exists(bundledScript))
        {
            return bundledScript;
        }

        throw new FileNotFoundException("Could not find the translation script.", bundledScript);
    }

    private static async Task<TranslationExecutionResult> RunTranslationAsync(
        string inputPath,
        string outputPath,
        string scriptPath,
        string sourceLanguage,
        string targetLanguage,
        string workingDirectory,
        CancellationToken cancellationToken)
    {
        var pythonCandidates = new[]
        {
            Environment.GetEnvironmentVariable("PPTX_TRANSLATOR_PYTHON"),
            "python",
            "py"
        }.Where(candidate => !string.IsNullOrWhiteSpace(candidate)).Cast<string>();

        var arguments = $"\"{scriptPath}\" \"{inputPath}\" \"{outputPath}\" --src {sourceLanguage} --dest {targetLanguage}";
        Exception? lastException = null;

        foreach (var pythonCandidate in pythonCandidates)
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = pythonCandidate,
                    Arguments = arguments,
                    WorkingDirectory = workingDirectory,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start the Python translation process.");
                var standardOutputTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
                var standardErrorTask = process.StandardError.ReadToEndAsync(cancellationToken);
                await process.WaitForExitAsync(cancellationToken);

                var standardOutput = await standardOutputTask;
                var standardError = await standardErrorTask;

                return new TranslationExecutionResult(
                    process.ExitCode,
                    standardOutput,
                    standardError,
                    pythonCandidate,
                    process.ExitCode == 0 && File.Exists(outputPath),
                    ExtractFailedChunkCount(standardOutput));
            }
            catch (Exception ex) when (ex is Win32Exception or InvalidOperationException)
            {
                lastException = ex;
            }
        }

        throw new InvalidOperationException("No usable Python interpreter was found. Set PPTX_TRANSLATOR_PYTHON if needed.", lastException);
    }

    private static int ExtractFailedChunkCount(string standardOutput)
    {
        var match = FailedChunksRegex().Match(standardOutput);
        return match.Success && int.TryParse(match.Groups[1].Value, out var failedChunkCount)
            ? failedChunkCount
            : 0;
    }

    [GeneratedRegex(@"Failed chunks kept in original language:\s*(\d+)", RegexOptions.IgnoreCase)]
    private static partial Regex FailedChunksRegex();

    private sealed record TranslationRequestFiles(string InputPath, string OutputPath, string OutputFileName);
    private sealed record TranslationExecutionResult(
        int ExitCode,
        string StandardOutput,
        string StandardError,
        string PythonCommand,
        bool Succeeded,
        int PartialFailureCount);

    private sealed record TranslationJob(
        string Id,
        TranslationJobStatus Status,
        string SourceLanguage,
        string TargetLanguage,
        string WorkingDirectory,
        string InputPath,
        string OutputPath,
        string OutputFileName,
        DateTimeOffset CreatedAtUtc,
        DateTimeOffset UpdatedAtUtc,
        DateTimeOffset ExpiresAtUtc,
        int PartialFailureCount,
        string? ErrorMessage,
        string? PythonCommand,
        string? StandardOutput,
        string? StandardError);

    private sealed record TranslationJobResponse(
        string JobId,
        string Status,
        string? Message,
        string? DownloadUrl,
        string? FileName,
        int? PartialFailureCount,
        DateTimeOffset CreatedAtUtc,
        DateTimeOffset UpdatedAtUtc,
        DateTimeOffset ExpiresAtUtc);

    private enum TranslationJobStatus
    {
        Queued,
        Running,
        Completed,
        Failed
    }

    private static string ToApiValue(this TranslationJobStatus status) => status switch
    {
        TranslationJobStatus.Queued => "queued",
        TranslationJobStatus.Running => "running",
        TranslationJobStatus.Completed => "completed",
        TranslationJobStatus.Failed => "failed",
        _ => "unknown"
    };
}
