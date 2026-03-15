using Aspire.Hosting.Docker;

var builder = DistributedApplication.CreateBuilder(args);

builder.AddDockerComposeEnvironment("env");

builder.AddDockerfile("webfrontend", "..", "PptxTranslator.Web/Dockerfile")
    .WithHttpEndpoint(targetPort: 8080)
    .WithEnvironment("ASPNETCORE_FORWARDEDHEADERS_ENABLED", "true")
    .WithEnvironment("ASPNETCORE_ENVIRONMENT", "Development");

builder.Build().Run();
