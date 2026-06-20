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
        await using var factory = CreateFactory();

        var client = factory.CreateClient(new WebApplicationFactoryClientOptions
        {
            AllowAutoRedirect = false
        });

        var response = await client.GetAsync("/openapi/v1.json");

        Assert.True(
            response.IsSuccessStatusCode,
            $"Expected success but got {(int)response.StatusCode} {response.StatusCode}");
    }

    [Fact]
    public async Task InvitationEndpoint_WithoutToken_ReturnsUnauthorized()
    {
        await using var factory = CreateFactory();

        var client = factory.CreateClient(new WebApplicationFactoryClientOptions
        {
            AllowAutoRedirect = false
        });

        var response = await client.PostAsJsonAsync("/api/fsdh/invitation", new
        {
            email = "someone@domain.com"
        });

        Assert.Equal(HttpStatusCode.Unauthorized, response.StatusCode);
    }

    private static WebApplicationFactory<global::Program> CreateFactory()
    {
        return new WebApplicationFactory<global::Program>()
            .WithWebHostBuilder(builder =>
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
            });
    }
}
