// <copyright file="ITeamPreferenceStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces
{
    using System.Collections.Generic;
    using Microsoft.Teams.Apps.SubmitIdea.Models;

    /// <summary>
    /// Interface for storage helper which helps in preparing model data for team preference.
    /// </summary>
    public interface ITeamPreferenceStorageHelper
    {
        /// <summary>
        /// Create team preference model data.
        /// </summary>
        /// <param name="entity">Team preference detail.</param>
        /// <returns>Represents team preference entity model.</returns>
        TeamPreferenceEntity CreateTeamPreferenceModel(TeamPreferenceEntity entity);

        /// <summary>
        /// Get posts unique tags.
        /// </summary>
        /// <param name="teamIdeas">Team post entities.</param>
        /// <param name="searchText">Input tag as search text.</param>
        /// <returns>Represents team tags.</returns>
        IEnumerable<string> GetUniqueTags(IEnumerable<TeamIdeaEntity> teamIdeas, string searchText);
    }
}
