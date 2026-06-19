using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Serialization;
using Microsoft.Kiota.Serialization.Json;
using Moq;

namespace InviterService.Tests;

public static class RequestAdapterMockHelper
{
    /// <summary>
    /// Creates a mock RequestAdapter configured to satisfy Microsoft Graph's
    /// internal serializing and URL routing requirements.
    /// </summary>
    public static Mock<IRequestAdapter> CreateMockAdapter()
    {
        var mockAdapter = new Mock<IRequestAdapter>();

        // Mock the serialization system so POST request bodies (like invitations) don't crash the SDK
        var mockSerializationFactory = new Mock<ISerializationWriterFactory>();
        mockSerializationFactory
            .Setup(f => f.GetSerializationWriter(It.IsAny<string>()))
            .Returns(() => new JsonSerializationWriter());

        mockAdapter.SetupGet(a => a.SerializationWriterFactory)
                    .Returns(mockSerializationFactory.Object);

        // Set a fake Base URL so the client doesn't throw null reference exceptions
        mockAdapter.SetupGet(a => a.BaseUrl).Returns("https://graph.microsoft.com/v1.0");
        mockAdapter.SetupSet(a => a.BaseUrl = It.IsAny<string>());

        return mockAdapter;
    }
}
