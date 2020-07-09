// <copyright file="MustBeValidUserHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Authentication
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc.Filters;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.SubmitIdea.Common;
    using Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces;
    using Microsoft.Teams.Apps.SubmitIdea.Models;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// This class is an authorization handler, which handles the authorization requirement.
    /// </summary>
    public class MustBeValidUserHandler : AuthorizationHandler<MustBeValidUserRequirement>
    {
        private readonly bool disableAuthentication;

        /// <summary>
        /// Bot adapter.
        /// </summary>
        private readonly IBotFrameworkHttpAdapter botAdapter;

        /// <summary>
        /// Provider to fetch team details from Azure Table Storage.
        /// </summary>
        private readonly ITeamStorageProvider teamStorageProvider;

        /// <summary>
        /// Microsoft application credentials.
        /// </summary>
        private readonly MicrosoftAppCredentials microsoftAppCredentials;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<MustBeValidUserHandler> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="MustBeValidUserHandler"/> class.
        /// </summary>
        /// <param name="configuration">ASP.NET Core <see cref="IConfiguration"/> instance.</param>
        public MustBeValidUserHandler(IConfiguration configuration)
        {
            this.disableAuthentication = configuration.GetValue<bool>("DisableAuthentication", false);
        }

        /// <summary>
        /// This method handles the authorization requirement.
        /// </summary>
        /// <param name="context">AuthorizationHandlerContext instance.</param>
        /// <param name="requirement">IAuthorizationRequirement instance.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected async override Task HandleRequirementAsync(
            AuthorizationHandlerContext context,
            MustBeValidUserRequirement requirement)
        {
            context = context ?? throw new ArgumentNullException(nameof(context));

            string teamId = string.Empty;
            var oidClaimType = "http://schemas.microsoft.com/identity/claims/objectidentifier";

            var oidClaim = context.User.Claims.FirstOrDefault(p => oidClaimType.Equals(p.Type, StringComparison.OrdinalIgnoreCase));

            if (context.Resource is AuthorizationFilterContext authorizationFilterContext)
            {
                // Wrap the request stream so that we can rewind it back to the start for regular request processing.
                authorizationFilterContext.HttpContext.Request.EnableBuffering();

                if (string.IsNullOrEmpty(authorizationFilterContext.HttpContext.Request.QueryString.Value))
                {
                    // Read the request body, parse out the activity object, and set the parsed culture information.
                    var streamReader = new StreamReader(authorizationFilterContext.HttpContext.Request.Body, Encoding.UTF8, true, 1024, leaveOpen: true);
                    using (var jsonReader = new JsonTextReader(streamReader))
                    {
                        var obj = JObject.Load(jsonReader);
                        var tagEntity = obj.ToObject<TeamTagEntity>();
                        authorizationFilterContext.HttpContext.Request.Body.Seek(0, SeekOrigin.Begin);
                        teamId = tagEntity.TeamId;
                    }
                }
                else
                {
                    var requestQuery = authorizationFilterContext.HttpContext.Request.Query;
                    teamId = requestQuery.Where(queryData => queryData.Key == "teamId").Select(queryData => queryData.Value.ToString()).FirstOrDefault();
                }
            }

            if (await this.ValidateUserAsync(teamId, oidClaim?.Value))
            {
                context.Succeed(requirement);
            }
        }

        /// <summary>
        /// Check if a user is a member of a certain team.
        /// </summary>
        /// <param name="teamId">The team id that the validator uses to check if the user is a member of the team. </param>
        /// <param name="userAadObjectId">The user's Azure Active Directory object id.</param>
        /// <returns>The flag indicates that the user is a part of certain team or not.</returns>
        private async Task<bool> ValidateUserAsync(string teamId, string userAadObjectId)
        {
            var teamMember = await this.GetTeamMemberAsync(teamId, userAadObjectId);
            return teamMember != null;
        }

        /// <summary>
        /// To fetch team member information for specified team.
        /// Return null if the member is not found in team id or either of the information is incorrect.
        /// Caller should handle null value to throw unauthorized if required
        /// </summary>
        /// <param name="teamId">Team id.</param>
        /// <param name="userId">User object id.</param>
        /// <returns>Returns team member information.</returns>
        private async Task<TeamsChannelAccount> GetTeamMemberAsync(string teamId, string userId)
        {
            TeamsChannelAccount teamMember = new TeamsChannelAccount();

            try
            {
                var teamDetails = await this.teamStorageProvider.GetTeamDetailAsync(teamId);
                string serviceUrl = teamDetails.ServiceUrl;

                var conversationReference = new ConversationReference
                {
                    ChannelId = Constants.TeamsBotFrameworkChannelId,
                    ServiceUrl = serviceUrl,
                };
                await ((BotFrameworkAdapter)this.botAdapter).ContinueConversationAsync(
                    this.microsoftAppCredentials.MicrosoftAppId,
                    conversationReference,
                    async (context, token) =>
                    {
                        teamMember = await TeamsInfo.GetTeamMemberAsync(context, userId, teamId, CancellationToken.None);
                    }, default);
            }
#pragma warning disable CA1031 // Catching general exceptions to log exception details in telemetry client.
            catch (Exception ex)
#pragma warning restore CA1031 // Catching general exceptions to log exception details in telemetry client.
            {
                this.logger.LogError(ex, $"Error occurred while fetching team member for team: {teamId} - user object id: {userId} ");

                // Return null if the member is not found in team id or either of the information is incorrect.
                // Caller should handle null value to throw unauthorized if required.
                return null;
            }

            return teamMember;
        }
    }
}
