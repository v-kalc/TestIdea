// <copyright file="ITeamCategoryStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.SubmitIdea.Models;

    /// <summary>
    /// Interface for provider which helps in storing, updating or deleting team categories in Microsoft Azure Table storage.
    /// </summary>
    public interface ITeamCategoryStorageProvider
    {
        /// <summary>
        /// Stores or update team categories data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamCategoryEntity">Holds team preference detail entity data.</param>
        /// <returns>A task that represents team preference entity data is saved or updated.</returns>
        Task<bool> UpsertTeamCategoriesAsync(TeamCategoryEntity teamCategoryEntity);

        /// <summary>
        /// Get team categories data from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamId">Team id for which need to fetch data.</param>
        /// <returns>A task that represents to hold team categories data.</returns>
        Task<TeamCategoryEntity> GetTeamCategoriesDataAsync(string teamId);

        /// <summary>
        /// Delete configured categories for a team if Bot is uninstalled from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamId">Holds team id.</param>
        /// <returns>A task that represents team categories data is deleted.</returns>
        Task<bool> DeleteTeamCategoriesEntryDataAsync(string teamId);
    }
}
