// <copyright file="TeamPreferenceController.cs" company="Microsoft">
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
    /// Controller to handle team preference API operations.
    /// </summary>
    [Route("api/teampreference")]
    [ApiController]
    [Authorize]
    public class TeamPreferenceController : BaseSubmitIdeaController
    {
        /// <summary>
        /// Event name for team preference HTTP get call.
        /// </summary>
        private const string RecordTeamPreferencePostHTTPGetCall = "Team preferences - HTTP Get call succeeded";

        /// <summary>
        /// Event name for team preference tags HTTP get call.
        /// </summary>
        private const string RecordTeamPreferenceTagsHTTPGetCall = "Team preferences tags - HTTP Get call succeeded";

        /// <summary>
        /// Event name for team preference HTTP post call.
        /// </summary>
        private const string RecordTeamPreferenceHTTPPostCall = "Team preferences - HTTP Post call succeeded";

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Instance of team preference storage helper.
        /// </summary>
        private readonly ITeamPreferenceStorageHelper teamPreferenceStorageHelper;

        /// <summary>
        /// Instance of team preference storage provider for team preferences.
        /// </summary>
        private readonly ITeamPreferenceStorageProvider teamPreferenceStorageProvider;

        /// <summary>
        /// Instance of Search service for working with Microsoft Azure Table storage.
        /// </summary>
        private readonly ITeamIdeaSearchService teamIdeaSearchService;

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamPreferenceController"/> class.
        /// </summary>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        /// <param name="teamPreferenceStorageHelper">Team preference storage helper dependency injection.</param>
        /// <param name="teamPreferenceStorageProvider">Team preference storage provider dependency injection.</param>
        /// <param name="teamIdeaSearchService">The team post search service dependency injection.</param>
        public TeamPreferenceController(
            ILogger<TeamPreferenceController> logger,
            TelemetryClient telemetryClient,
            ITeamPreferenceStorageHelper teamPreferenceStorageHelper,
            ITeamPreferenceStorageProvider teamPreferenceStorageProvider,
            ITeamIdeaSearchService teamIdeaSearchService)
            : base(telemetryClient)
        {
            this.logger = logger;
            this.teamPreferenceStorageHelper = teamPreferenceStorageHelper;
            this.teamPreferenceStorageProvider = teamPreferenceStorageProvider;
            this.teamIdeaSearchService = teamIdeaSearchService;
        }

        /// <summary>
        /// Get call to retrieve team preference data.
        /// </summary>
        /// <param name="teamId">Team id - unique value for each Team where preference has configured.</param>
        /// <returns>Represents Team preference entity model.</returns>
        [HttpGet]
        public async Task<IActionResult> GetAsync(string teamId)
        {
            try
            {
                this.logger.LogInformation("Call to retrieve list of team preference.");

                if (string.IsNullOrEmpty(teamId))
                {
                    this.logger.LogError("Error while getting the team preference from Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while getting team preference details from Microsoft Azure Table storage.");
                }

                var teamPreference = await this.teamPreferenceStorageProvider.GetTeamPreferenceDataAsync(teamId);
                this.RecordEvent(RecordTeamPreferencePostHTTPGetCall);

                return this.Ok(teamPreference);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to team preference service.");
                throw;
            }
        }

        /// <summary>
        /// Post call to store team preference details in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamPreferenceEntity">Holds team preference detail entity data.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost]
        public async Task<IActionResult> PostAsync([FromBody] TeamPreferenceEntity teamPreferenceEntity)
        {
            try
            {
                this.logger.LogInformation("Call to add team preference.");

                if (string.IsNullOrEmpty(teamPreferenceEntity?.TeamId))
                {
                    this.logger.LogError("Error while creating or updating team preference details data in Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while creating or updating team preference details in Microsoft Azure Table storage.");
                }

                this.RecordEvent(RecordTeamPreferenceHTTPPostCall);
                var teamPreferenceDetail = this.teamPreferenceStorageHelper.CreateTeamPreferenceModel(teamPreferenceEntity);

                return this.Ok(await this.teamPreferenceStorageProvider.UpsertTeamPreferenceAsync(teamPreferenceDetail));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to team preference service.");
                throw;
            }
        }

        /// <summary>
        /// Get list of unique tags to show while configuring the preference.
        /// </summary>
        /// <param name="searchText">Search text represents the text to find and get unique tags.</param>
        /// <returns>List of unique tags.</returns>
        [HttpGet("unique-tags")]
        public async Task<IActionResult> GetUniqueTagsAsync(string searchText)
        {
            try
            {
                this.logger.LogInformation("Call to get list of unique tags to show while configuring the preference.");

                if (string.IsNullOrEmpty(searchText))
                {
                    this.logger.LogError("Error while getting the list of unique tags from Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while getting the list of unique tags from Microsoft Azure Table storage.");
                }

                var teamPosts = await this.teamIdeaSearchService.GetTeamIdeasAsync(TeamPostSearchScope.TeamPreferenceTags, searchText, userObjectId: null);
                var uniqueTags = this.teamPreferenceStorageHelper.GetUniqueTags(teamPosts, searchText);
                this.RecordEvent(RecordTeamPreferenceTagsHTTPGetCall);

                return this.Ok(uniqueTags);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to get unique tags.");
                throw;
            }
        }
    }
}