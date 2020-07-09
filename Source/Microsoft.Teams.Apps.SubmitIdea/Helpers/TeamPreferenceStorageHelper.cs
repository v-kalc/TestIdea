// <copyright file="TeamPreferenceStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces;
    using Microsoft.Teams.Apps.SubmitIdea.Models;

    /// <summary>
    /// Implements team preference storage helper which helps to construct the model, get unique tags for team preference.
    /// </summary>
    public class TeamPreferenceStorageHelper : ITeamPreferenceStorageHelper
    {
        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<TeamPreferenceStorageHelper> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamPreferenceStorageHelper"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public TeamPreferenceStorageHelper(
            ILogger<TeamPreferenceStorageHelper> logger)
        {
            this.logger = logger;
        }

        /// <summary>
        /// Create team preference model data.
        /// </summary>
        /// <param name="entity">Represents team preference entity object.</param>
        /// <returns>Represents team preference entity model.</returns>
        public TeamPreferenceEntity CreateTeamPreferenceModel(TeamPreferenceEntity entity)
        {
            try
            {
                entity = entity ?? throw new ArgumentNullException(nameof(entity));

                entity.PreferenceId = Guid.NewGuid().ToString();
                entity.CreatedDate = DateTime.UtcNow;
                entity.UpdatedDate = DateTime.UtcNow;

                return entity;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception occurred while preparing the team preference entity model data");
                throw;
            }
        }

        /// <summary>
        /// Get posts unique tags.
        /// </summary>
        /// <param name="teamIdeas">Team post entities.</param>
        /// <param name="searchText">Search text for tags.</param>
        /// <returns>Represents team tags.</returns>
        public IEnumerable<string> GetUniqueTags(IEnumerable<TeamIdeaEntity> teamIdeas, string searchText)
        {
            try
            {
                teamIdeas = teamIdeas ?? throw new ArgumentNullException(nameof(teamIdeas));
                var tags = new List<string>();

                if (searchText == "*")
                {
                    foreach (var teamPost in teamIdeas)
                    {
                        tags.AddRange(teamPost.Tags?.Split(";"));
                    }
                }
                else
                {
                    foreach (var teamPost in teamIdeas)
                    {
                        tags.AddRange(teamPost.Tags?.Split(";").Where(tag => tag.Contains(searchText, StringComparison.InvariantCultureIgnoreCase)));
                    }
                }

                return tags.Distinct().OrderBy(tag => tag);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception occurred while preparing the team preference entity model data");
                throw;
            }
        }
    }
}
