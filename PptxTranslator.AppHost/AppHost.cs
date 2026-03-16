using Aspire.Hosting.Docker;

var builder = DistributedApplication.CreateBuilder(args);

builder.AddDockerComposeEnvironment("env");

builder.AddDockerfile("webfrontend", "..", "PptxTranslator.Web/Dockerfile")
    .WithHttpEndpoint(port: 5000, targetPort: 8080)
    .WithExternalHttpEndpoints()
    .WithEnvironment("ASPNETCORE_FORWARDEDHEADERS_ENABLED", "true");

builder.Build().Run();
