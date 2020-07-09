// <copyright file="ITeamTagStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.SubmitIdea.Models;

    /// <summary>
    /// Interface for provider which helps in storing, updating or deleting team tags in Microsoft Azure Table storage.
    /// </summary>
    public interface ITeamTagStorageProvider
    {
        /// <summary>
        /// Stores or update team tags data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamTagEntity">Holds team preference detail entity data.</param>
        /// <returns>A task that represents team preference entity data is saved or updated.</returns>
        Task<bool> UpsertTeamTagsAsync(TeamTagEntity teamTagEntity);

        /// <summary>
        /// Get team tags data from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamId">Team id for which need to fetch data.</param>
        /// <returns>A task that represents to hold team tags data.</returns>
        Task<TeamTagEntity> GetTeamTagsDataAsync(string teamId);

        /// <summary>
        /// Delete configured tags for a team if Bot is uninstalled from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamId">Holds team id.</param>
        /// <returns>A task that represents team tags data is deleted.</returns>
        Task<bool> DeleteTeamTagsEntryDataAsync(string teamId);
    }
}
