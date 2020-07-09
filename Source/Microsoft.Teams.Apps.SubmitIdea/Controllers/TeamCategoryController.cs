// <copyright file="TeamCategoryController.cs" company="Microsoft">
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
    /// Controller to handle team categories API operations.
    /// </summary>
    [Route("api/teamcategory")]
    [ApiController]
    [Authorize]
    public class TeamCategoryController : BaseSubmitIdeaController
    {
        /// <summary>
        /// Event name for team category HTTP get call.
        /// </summary>
        private const string RecordTeamTagHTTPGetCall = "Team categories - HTTP Get call succeeded";

        /// <summary>
        /// Event name for team category HTTP get call.
        /// </summary>
        private const string RecordTeamConfiguredTagsHTTPGetCall = "Team categories - HTTP Get call succeeded";

        /// <summary>
        /// Event name for team category HTTP post call.
        /// </summary>
        private const string RecordTeamTagHTTPPostCall = "Team categories - HTTP Post call succeeded";

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Instance of team category storage helper.
        /// </summary>
        private readonly ITeamCategoryStorageHelper teamCategoryStorageHelper;

        /// <summary>
        /// Instance of team category storage provider for team categories.
        /// </summary>
        private readonly ITeamCategoryStorageProvider teamCategoryStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamCategoryController"/> class.
        /// </summary>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        /// <param name="teamCategoryStorageHelper">Team category storage helper dependency injection.</param>
        /// <param name="teamCategoryStorageProvider">Team category storage provider dependency injection.</param>
        public TeamCategoryController(
            ILogger<TeamCategoryController> logger,
            TelemetryClient telemetryClient,
            ITeamCategoryStorageHelper teamCategoryStorageHelper,
            ITeamCategoryStorageProvider teamCategoryStorageProvider)
            : base(telemetryClient)
        {
            this.logger = logger;
            this.teamCategoryStorageHelper = teamCategoryStorageHelper;
            this.teamCategoryStorageProvider = teamCategoryStorageProvider;
        }

        /// <summary>
        /// Get call to retrieve team categories data.
        /// </summary>
        /// <param name="teamId">Team Id - unique value for each Team where categories has configured.</param>
        /// <returns>Represents Team category entity model.</returns>
        [HttpGet]
        public async Task<IActionResult> GetAsync(string teamId)
        {
            try
            {
                this.logger.LogInformation("Call to retrieve team categories data.");

                if (string.IsNullOrEmpty(teamId))
                {
                    this.logger.LogError("Error while getting the team categories from Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while getting the team categories from Microsoft Azure Table storage.");
                }

                var teamPreference = await this.teamCategoryStorageProvider.GetTeamCategoriesDataAsync(teamId);
                this.RecordEvent(RecordTeamTagHTTPGetCall);

                return this.Ok(teamPreference);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to team categories service.");
                throw;
            }
        }

        /// <summary>
        /// Post call to store team category details in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamCategoryEntity">Holds team category detail entity data.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost]
        public async Task<IActionResult> PostAsync([FromBody] TeamCategoryEntity teamCategoryEntity)
        {
            try
            {
                this.logger.LogInformation("Call to add team category details.");

                if (string.IsNullOrEmpty(teamCategoryEntity?.TeamId))
                {
                    this.logger.LogError("Error while creating or updating team category details in Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while creating or updating team category details in Microsoft Azure Table storage.");
                }

                this.RecordEvent(RecordTeamTagHTTPPostCall);
                var teamTagDetail = this.teamCategoryStorageHelper.CreateTeamCategoryModel(teamCategoryEntity, this.UserName, this.UserAadId);

                return this.Ok(await this.teamCategoryStorageProvider.UpsertTeamCategoriesAsync(teamTagDetail));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to team category service.");
                throw;
            }
        }

        /// <summary>
        /// Get list of configured categories for a team to show on filter bar dropdown list.
        /// </summary>
        /// <param name="teamId">Team id to get the configured categories for a team.</param>
        /// <returns>List of configured categories.</returns>
        [HttpGet("configured-categories")]
        public async Task<IActionResult> GetConfiguredTagsAsync(string teamId)
        {
            try
            {
                this.logger.LogInformation("Call to get list of configured categories for a team.");

                if (string.IsNullOrEmpty(teamId))
                {
                    this.logger.LogError("Error while getting the list of team categories from Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while getting the list of team categories from Microsoft Azure Table storage.");
                }

                var teamTagDetail = await this.teamCategoryStorageProvider.GetTeamCategoriesDataAsync(teamId);
                this.RecordEvent(RecordTeamConfiguredTagsHTTPGetCall);

                return this.Ok(teamTagDetail?.Categories?.Split(";"));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to get configured categories.");
                throw;
            }
        }
    }
}