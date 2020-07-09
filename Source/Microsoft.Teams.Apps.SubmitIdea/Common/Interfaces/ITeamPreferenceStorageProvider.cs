// <copyright file="ITeamPreferenceStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.SubmitIdea.Models;

    /// <summary>
    /// Interface for provider which helps in storing, updating team preference for posts in Microsoft Azure Table storage.
    /// </summary>
    public interface ITeamPreferenceStorageProvider
    {
        /// <summary>
        /// Stores or update team preference data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamPreferenceEntity">Holds team preference detail entity data.</param>
        /// <returns>A task that represents team preference entity data is saved or updated.</returns>
        Task<bool> UpsertTeamPreferenceAsync(TeamPreferenceEntity teamPreferenceEntity);

        /// <summary>
        /// Get team preference data from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamId">Team Id for which need to fetch data.</param>
        /// <returns>A task that represents to hold team preference data.</returns>
        Task<TeamPreferenceEntity> GetTeamPreferenceDataAsync(string teamId);

        /// <summary>
        /// Get team preferences data from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="digestFrequency">Digest frequency text for notification like Monthly/Weekly.</param>
        /// <returns>A task that represent collection to hold team preferences data.</returns>
        Task<IEnumerable<TeamPreferenceEntity>> GetTeamPreferencesAsync(string digestFrequency);
    }
}
