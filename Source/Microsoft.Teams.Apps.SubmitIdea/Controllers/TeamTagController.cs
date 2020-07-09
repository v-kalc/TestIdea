// <copyright file="TeamTagController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Controllers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces;
    using Microsoft.Teams.Apps.SubmitIdea.Models;

    /// <summary>
    /// Controller to handle team tags API operations.
    /// </summary>
    [Route("api/teamtag")]
    [ApiController]
    [Authorize]
    public class TeamTagController : BaseSubmitIdeaController
    {
        /// <summary>
        /// Event name for team tag HTTP get call.
        /// </summary>
        private const string RecordTeamTagHTTPGetCall = "Team tags - HTTP Get call succeeded";

        /// <summary>
        /// Event name for team tag HTTP get call.
        /// </summary>
        private const string RecordTeamConfiguredTagsHTTPGetCall = "Team tags - HTTP Get call succeeded";

        /// <summary>
        /// Event name for team tag HTTP post call.
        /// </summary>
        private const string RecordTeamTagHTTPPostCall = "Team tags - HTTP Post call succeeded";

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Instance of team tag storage helper.
        /// </summary>
        private readonly ITeamTagStorageHelper teamTagStorageHelper;

        /// <summary>
        /// Instance of team tag storage provider for team tags.
        /// </summary>
        private readonly ITeamTagStorageProvider teamTagStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamTagController"/> class.
        /// </summary>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        /// <param name="teamTagStorageHelper">Team tag storage helper dependency injection.</param>
        /// <param name="teamTagStorageProvider">Team tag storage provider dependency injection.</param>
        public TeamTagController(
            ILogger<TeamTagController> logger,
            TelemetryClient telemetryClient,
            ITeamTagStorageHelper teamTagStorageHelper,
            ITeamTagStorageProvider teamTagStorageProvider)
            : base(telemetryClient)
        {
            this.logger = logger;
            this.teamTagStorageHelper = teamTagStorageHelper;
            this.teamTagStorageProvider = teamTagStorageProvider;
        }

        /// <summary>
        /// Get call to retrieve team tags data.
        /// </summary>
        /// <param name="teamId">Team Id - unique value for each Team where tags has configured.</param>
        /// <returns>Represents Team tag entity model.</returns>
        [HttpGet]
        public async Task<IActionResult> GetAsync(string teamId)
        {
            try
            {
                this.logger.LogInformation("Call to retrieve team tags data.");

                if (string.IsNullOrEmpty(teamId))
                {
                    this.logger.LogError("Error while getting the team tags from Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while getting the team tags from Microsoft Azure Table storage.");
                }

                var teamPreference = await this.teamTagStorageProvider.GetTeamTagsDataAsync(teamId);
                this.RecordEvent(RecordTeamTagHTTPGetCall);

                return this.Ok(teamPreference);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to team tags service.");
                throw;
            }
        }

        /// <summary>
        /// Post call to store team tag details in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamTagEntity">Holds team tag detail entity data.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost]
        public async Task<IActionResult> PostAsync([FromBody] TeamTagEntity teamTagEntity)
        {
            try
            {
                this.logger.LogInformation("Call to add team tag details.");
                teamTagEntity = teamTagEntity ?? throw new ArgumentNullException(nameof(teamTagEntity));

                if (string.IsNullOrEmpty(teamTagEntity.TeamId))
                {
                    this.logger.LogError("Error while creating or updating team tag details in Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while creating or updating team tag details in Microsoft Azure Table storage.");
                }

                this.RecordEvent(RecordTeamTagHTTPPostCall);
                var teamTagDetail = new TeamTagEntity
                {
                    CreatedByName = this.UserName,
                    UserAadId = this.UserAadId,
                    CreatedDate = DateTime.UtcNow,
                    Tags = teamTagEntity.Tags,
                    TeamId = teamTagEntity.TeamId,
                };

                return this.Ok(await this.teamTagStorageProvider.UpsertTeamTagsAsync(teamTagDetail));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to team tag service.");
                throw;
            }
        }

        /// <summary>
        /// Get list of configured tags for a team to show on filter bar dropdown list.
        /// </summary>
        /// <param name="teamId">Team id to get the configured tags for a team.</param>
        /// <returns>List of configured tags.</returns>
        [HttpGet("configured-tags")]
        public async Task<IActionResult> GetConfiguredTagsAsync(string teamId)
        {
            try
            {
                this.logger.LogInformation("Call to get list of configured tags for a team.");

                if (string.IsNullOrEmpty(teamId))
                {
                    this.logger.LogError("Error while getting the list of team tags from Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while getting the list of team tags from Microsoft Azure Table storage.");
                }

                var teamTagDetail = await this.teamTagStorageProvider.GetTeamTagsDataAsync(teamId);
                this.RecordEvent(RecordTeamConfiguredTagsHTTPGetCall);

                return this.Ok(teamTagDetail?.Tags?.Split(";"));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to get configured tags.");
                throw;
            }
        }
    }
}