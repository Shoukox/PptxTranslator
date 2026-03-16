using System.ComponentModel;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.Mvc;

namespace PptxTranslator.Web.Features.Translation;

internal static partial class TranslateFeature
{
    public static IEndpointRouteBuilder MapTranslateFeature(this IEndpointRouteBuilder endpoints)
    {
        endpoints.MapPost("/translate", HandleTranslateAsync);
        return endpoints;
    }

    [RequestSizeLimit(1024 * 1024 * 1024)] // 1 GB
    [RequestFormLimits(MultipartBodyLengthLimit = 1024 * 1024 * 1024)] // 1 GB
    private static async Task<IResult> HandleTranslateAsync(
        HttpContext httpContext,
        HttpRequest request,
        IWebHostEnvironment environment,
        ILogger<Program> logger,
        CancellationToken cancellationToken)
    {
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

        try
        {
            var requestFiles = await SaveRequestFilesAsync(file, targetLanguage, workingDirectory, cancellationToken);
            var scriptPath = ResolveScriptPath(environment.ContentRootPath);

            var translationResult = await RunTranslationAsync(
                requestFiles.InputPath,
                requestFiles.OutputPath,
                scriptPath,
                sourceLanguage,
                targetLanguage,
                workingDirectory,
                cancellationToken);

            if (!translationResult.Succeeded)
            {
                logger.LogError(
                    "Translation failed. ExitCode: {ExitCode}, Python: {PythonCommand}, StdOut: {StdOut}, StdErr: {StdErr}",
                    translationResult.ExitCode,
                    translationResult.PythonCommand,
                    translationResult.StandardOutput,
                    translationResult.StandardError);

                return Results.Problem(
                    title: "Translation failed",
                    detail: "Something went wrong...",
                    statusCode: StatusCodes.Status500InternalServerError);
            }
            logger.LogInformation(
                "Translation succeeded. Python: {PythonCommand}, StdOut: {StdOut}, StdErr: {StdErr}",
                translationResult.PythonCommand,
                translationResult.StandardOutput,
                translationResult.StandardError);

            if (translationResult.PartialFailureCount > 0)
            {
                httpContext.Response.Headers.Append(
                    "X-Translation-Notice",
                    "Translation completed. Some text was kept in the original language.");
            }

            var bytes = await File.ReadAllBytesAsync(requestFiles.OutputPath, cancellationToken);
            return Results.File(
                bytes,
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                requestFiles.OutputFileName);
        }
        finally
        {
            DeleteWorkingDirectory(workingDirectory, logger);
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
}
