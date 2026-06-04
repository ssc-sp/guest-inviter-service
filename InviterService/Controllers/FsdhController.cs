using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace ApiBackend.Controllers
{
    public class FsdhInviteRequest
    {
        public string email { get; set; } = string.Empty;
    }

    public class FsdhInviteResponse
    {
        public string objectid { get; set; } = string.Empty;
        public string? redeemUrl { get; set; }
    }

    [Route("api/fsdh")]
    [ApiController]
    //[Authorize(Roles = "fsdh_api_access,internal_api_access")]
    public class FsdhController : ControllerBase
    {
        private readonly GraphServiceClient graphServiceClient;

        public FsdhController(IConfiguration config, GraphServiceClient graphServiceClient)
        {
            this.graphServiceClient = graphServiceClient;
        }

        [HttpPost("invitation")]
        [Authorize]
        [ProducesResponseType(statusCode: StatusCodes.Status200OK, type: typeof(FsdhInviteResponse))]
        [ProducesResponseType(statusCode: StatusCodes.Status201Created, type: typeof(FsdhInviteResponse))]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public async Task<IActionResult> CreateInvitation([FromBody] FsdhInviteRequest request)
        {
            if (string.IsNullOrEmpty(request.email))
            {
                return BadRequest("Email required");
            }

            var response = new FsdhInviteResponse();

            try
            {
                var userLookupCollection = await graphServiceClient.Users.GetAsync(query =>
                {
                    query.QueryParameters.Filter = $"mail eq '{request.email}'";
                    query.QueryParameters.Select = ["id"];
                }) ?? throw new Exception("Unexpected null user lookup response");

                var results = PageIterator<User, UserCollectionResponse>.CreatePageIterator(graphServiceClient, userLookupCollection, user => false);

                var userId = userLookupCollection?.Value?.FirstOrDefault()?.Id;

                if (userId == null)
                {
                    var invite = await graphServiceClient.Invitations.PostAsync(new Invitation
                    {
                        InvitedUserEmailAddress = request.email,
                        InviteRedirectUrl = "https://portal.azure.com", // This URL should probably be the FSDH application url so they can login after accepting the invite
                        SendInvitationMessage = false // This allows the FSDH application to use a custom email template
                    });

                    // Checking for unexpected nulls
                    if (invite == null || invite.InvitedUser == null || invite.InvitedUser.Id == null)
                    {
                        throw new Exception("Invitation failed");
                    }

                    response.objectid = invite.InvitedUser.Id;
                    response.redeemUrl = invite.InviteRedeemUrl;
                }
                else
                {
                    response.objectid = userId;
                }

                return userId == null ? Created("", response) : Ok(response);
            }
            catch (Exception)
            {                
                return StatusCode(StatusCodes.Status500InternalServerError, "Error creating invitation");
            }
        }

    }
}
