using ApiBackend.Controllers;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Serialization;
using Moq;
using Xunit;

namespace InviterService.Tests;

public class FsdhControllerTests
{
    private readonly Mock<IConfiguration> _mockConfig;

    public FsdhControllerTests()
    {
        _mockConfig = new Mock<IConfiguration>();
    }

    [Fact]
    public async Task CreateInvitation_WhenEmailIsEmpty_ReturnsBadRequest()
    {
        // Arrange
        var controller = new FsdhController(_mockConfig.Object, null!);
        var request = new FsdhInviteRequest { email = "" };

        // Act
        var result = await controller.CreateInvitation(request);

        // Assert
        var badRequestResult = Assert.IsType<BadRequestObjectResult>(result);
        Assert.Equal("Email required", badRequestResult.Value);
    }

    [Fact]
    public async Task CreateInvitation_WhenUserAlreadyExists_ReturnsOkWithExistingId()
    {
        // Arrange
        var testEmail = "existing@domain.com";
        var existingUserId = "existing-user-guid-123";

        var userListResponse = new UserCollectionResponse
        {
            Value = new List<User> { new() { Id = existingUserId, Mail = testEmail } }
        };

        var mockAdapter = RequestAdapterMockHelper.CreateMockAdapter();
        
        mockAdapter.Setup(a => a.SendAsync(
            It.IsAny<RequestInformation>(),
            It.IsAny<ParsableFactory<UserCollectionResponse>>(),
            It.IsAny<Dictionary<string, ParsableFactory<Microsoft.Kiota.Abstractions.Serialization.IParsable>>>(),
            It.IsAny<CancellationToken>()))
        .ReturnsAsync(userListResponse);

        var graphServiceClient = new GraphServiceClient(mockAdapter.Object);
        var controller = new FsdhController(_mockConfig.Object, graphServiceClient);
        var request = new FsdhInviteRequest { email = testEmail };

        // Act
        var result = await controller.CreateInvitation(request);

        // Assert
        var okResult = Assert.IsType<OkObjectResult>(result);
        var responseBody = Assert.IsType<FsdhInviteResponse>(okResult.Value);
        Assert.Equal(existingUserId, responseBody.objectid);
        Assert.Null(responseBody.redeemUrl);
    }

    [Fact]
    public async Task CreateInvitation_WhenUserIsNew_InvitesUserAndReturnsCreated()
    {
        // Arrange
        var testEmail = "new-user@domain.com";
        var newUserId = "newly-created-guid-456";
        var mockRedeemUrl = "https://login.microsoftonline.com/redeem?tc=xyz";

        var mockAdapter = RequestAdapterMockHelper.CreateMockAdapter();

        var emptyUserResponse = new UserCollectionResponse { Value = new List<User>() };
        mockAdapter.Setup(a => a.SendAsync(
            It.IsAny<RequestInformation>(),
            It.IsAny<ParsableFactory<UserCollectionResponse>>(),
            It.IsAny<Dictionary<string, ParsableFactory<Microsoft.Kiota.Abstractions.Serialization.IParsable>>>(),
            It.IsAny<CancellationToken>()))
        .ReturnsAsync(emptyUserResponse);

        var mockInvitationResponse = new Invitation
        {
            InvitedUser = new User { Id = newUserId }, // Fixed DirectoryObject type conversion mismatch
            InviteRedeemUrl = mockRedeemUrl
        };
        mockAdapter.Setup(a => a.SendAsync(
            It.IsAny<RequestInformation>(),
            It.IsAny<ParsableFactory<Invitation>>(),
            It.IsAny<Dictionary<string, ParsableFactory<Microsoft.Kiota.Abstractions.Serialization.IParsable>>>(),
            It.IsAny<CancellationToken>()))
        .ReturnsAsync(mockInvitationResponse);

        var graphServiceClient = new GraphServiceClient(mockAdapter.Object);
        var controller = new FsdhController(_mockConfig.Object, graphServiceClient);
        var request = new FsdhInviteRequest { email = testEmail };

        // Act
        var result = await controller.CreateInvitation(request);

        // Assert
        var createdResult = Assert.IsType<CreatedResult>(result);
        var responseBody = Assert.IsType<FsdhInviteResponse>(createdResult.Value);
        Assert.Equal(newUserId, responseBody.objectid);
        Assert.Equal(mockRedeemUrl, responseBody.redeemUrl);
    }

    [Fact]
    public async Task CreateInvitation_WhenGraphThrowsException_Returns500InternalServerError()
    {
        // Arrange
        var mockAdapter = RequestAdapterMockHelper.CreateMockAdapter();
        
        mockAdapter.Setup(a => a.SendAsync(
            It.IsAny<RequestInformation>(),
            It.IsAny<ParsableFactory<UserCollectionResponse>>(),
            It.IsAny<Dictionary<string, ParsableFactory<Microsoft.Kiota.Abstractions.Serialization.IParsable>>>(),
            It.IsAny<CancellationToken>()))
        .ThrowsAsync(new Exception("Simulated connection failure"));

        var graphServiceClient = new GraphServiceClient(mockAdapter.Object);
        var controller = new FsdhController(_mockConfig.Object, graphServiceClient);
        var request = new FsdhInviteRequest { email = "error@domain.com" };

        // Act
        var result = await controller.CreateInvitation(request);

        // Assert
        var statusCodeResult = Assert.IsType<ObjectResult>(result);
        Assert.Equal(StatusCodes.Status500InternalServerError, statusCodeResult.StatusCode);
        Assert.Equal("Error creating invitation", statusCodeResult.Value);
    }
}