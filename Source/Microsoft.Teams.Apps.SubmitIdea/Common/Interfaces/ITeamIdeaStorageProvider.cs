// <copyright file="ITeamIdeaStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.SubmitIdea.Models;

    /// <summary>
    /// Interface for provider which helps in retrieving, storing, updating and deleting team idea details in Microsoft Azure Table storage.
    /// </summary>
    public interface ITeamIdeaStorageProvider
    {
        /// <summary>
        /// Stores or update team idea details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamIdeaEntity">Holds team idea detail entity data.</param>
        /// <returns>A boolean that represents team idea entity data is successfully saved/updated or not.</returns>
        Task<bool> UpsertTeamIdeaAsync(TeamIdeaEntity teamIdeaEntity);

        /// <summary>
        /// Get team idea data from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="ideaId">Idea id to fetch the post details.</param>
        /// <returns>A task that represent a object to hold team post data.</returns>
        Task<TeamIdeaEntity> GetTeamIdeaEntityAsync(string ideaId);

        /// <summary>
        /// Get post data.
        /// </summary>
        /// <param name="postCreatedByuserId">User id to fetch the post details.</param>
        /// <param name="postId">Post id to fetch the post details.</param>
        /// <returns>A task that represent a object to hold post data.</returns>
        Task<TeamIdeaEntity> GetPostAsync(string postCreatedByuserId, string postId);
    }
}
