// <copyright file="TeamCategoryStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Helpers
{
    using System;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces;
    using Microsoft.Teams.Apps.SubmitIdea.Models;

    /// <summary>
    ///  Implements team category storage helper which helps to construct the model for team category.
    /// </summary>
    public class TeamCategoryStorageHelper : ITeamCategoryStorageHelper
    {
        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<TeamCategoryStorageHelper> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamCategoryStorageHelper"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public TeamCategoryStorageHelper(
            ILogger<TeamCategoryStorageHelper> logger)
        {
            this.logger = logger;
        }

        /// <summary>
        /// Create team categories mode data.
        /// </summary>
        /// <param name="teamCategoryEntity">Represents team category entity object.</param>
        /// <param name="userName">User name who has configured the categories in team.</param>
        /// <param name="userAadId">Azure Active Directory id of the user.</param>
        /// <returns>Represents team categories entity model.</returns>
        public TeamCategoryEntity CreateTeamCategoryModel(TeamCategoryEntity teamCategoryEntity, string userName, string userAadId)
        {
            try
            {
                teamCategoryEntity = teamCategoryEntity ?? throw new ArgumentNullException(nameof(teamCategoryEntity));

                teamCategoryEntity.CreatedByName = userName;
                teamCategoryEntity.UserAadId = userAadId;
                teamCategoryEntity.CreatedDate = DateTime.UtcNow;

                return teamCategoryEntity;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception occurred while preparing the team categories entity model data");
                throw;
            }
        }
    }
}
