using System.Net;
using System.Net.Http.Json;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.Extensions.Configuration;
using Xunit;

namespace InviterService.Tests;

public class ProgramTests
{
    [Fact]
    public async Task OpenApiEndpoint_WhenDevelopment_ReturnsSuccess()
    {
        // Arrange
        await using var factory = new TestProgramWebApplicationFactory();

        var client = factory.CreateClient(new WebApplicationFactoryClientOptions
        {
            AllowAutoRedirect = false
        });

        // Act
        var response = await client.GetAsync("/openapi/v1.json");

        // Assert
        Assert.True(
            response.IsSuccessStatusCode,
            $"Expected success but got {(int)response.StatusCode} {response.StatusCode}");
    }

    [Fact]
    public async Task InvitationEndpoint_WithoutToken_ReturnsUnauthorized()
    {
        // Arrange
        await using var factory = new TestProgramWebApplicationFactory();

        var client = factory.CreateClient(new WebApplicationFactoryClientOptions
        {
            AllowAutoRedirect = false
        });

        var requestBody = new
        {
            email = "someone@domain.com"
        };

        // Act
        var response = await client.PostAsJsonAsync("/api/fsdh/invitation", requestBody);

        // Assert
        Assert.Equal(HttpStatusCode.Unauthorized, response.StatusCode);
    }

    // Creates a test host for Program.cs using the Development environment.
    // The fake AzureAd settings let the app build its authentication services
    // without needing real tenant/client values during tests.
    private sealed class TestProgramWebApplicationFactory : WebApplicationFactory<global::Program>
    {
        protected override void ConfigureWebHost(IWebHostBuilder builder)
        {
            builder.UseEnvironment("Development");

            builder.ConfigureAppConfiguration((_, config) =>
            {
                config.AddInMemoryCollection(new Dictionary<string, string?>
                {
                    ["AzureAd:Instance"] = "https://login.microsoftonline.com/",
                    ["AzureAd:TenantId"] = "11111111-1111-1111-1111-111111111111",
                    ["AzureAd:ClientId"] = "22222222-2222-2222-2222-222222222222"
                });
            });
        }
    }
}
