// <copyright file="ITeamIdeaStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Teams.Apps.SubmitIdea.Models;

    /// <summary>
    /// Interface for storage helper which helps in preparing model data for team post.
    /// </summary>
    public interface ITeamIdeaStorageHelper
    {
        /// <summary>
        /// Create team idea details model.
        /// </summary>
        /// <param name="teamIdeaEntity">Team idea object.</param>
        /// <param name="userId">Azure Active directory id of user.</param>
        /// <param name="userName">Author who created the idea.</param>
        /// <returns>A task that represents team idea entity data.</returns>
        TeamIdeaEntity CreateTeamIdeaModel(TeamIdeaEntity teamIdeaEntity, string userId, string userName);

        /// <summary>
        /// Create updated team idea model to save in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamIdeaEntity">Team idea detail.</param>
        /// <returns>A task that represents team idea entity updated data.</returns>
        TeamIdeaEntity CreateUpdatedTeamIdeaModel(TeamIdeaEntity teamIdeaEntity);

        /// <summary>
        /// Get filtered team ideas as per the configured tags.
        /// </summary>
        /// <param name="teamIdeas">Team idea entities.</param>
        /// <param name="searchText">Search text for tags.</param>
        /// <returns>Represents team ideas.</returns>
        IEnumerable<TeamIdeaEntity> GetFilteredTeamIdeasAsPerTags(IEnumerable<TeamIdeaEntity> teamIdeas, string searchText);

        /// <summary>
        /// Get tags query to fetch team ideas as per the configured tags.
        /// </summary>
        /// <param name="tags">Tags of a configured team idea.</param>
        /// <returns>Represents tags query to fetch team ideas.</returns>
        string GetTags(string tags);

        /// <summary>
        /// Get filtered team ideas as per the date range from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamIdeas">Team ideas data.</param>
        /// <param name="fromDate">Start date from which data should fetch.</param>
        /// <param name="toDate">End date till when data should fetch.</param>
        /// <returns>A task that represent collection to hold team ideas data.</returns>
        IEnumerable<TeamIdeaEntity> GetTeamIdeasInDateRangeAsync(IEnumerable<TeamIdeaEntity> teamIdeas, DateTime fromDate, DateTime toDate);

        /// <summary>
        /// Get filtered unique user names.
        /// </summary>
        /// <param name="teamIdeas">Team idea entities.</param>
        /// <returns>Represents team ideas.</returns>
        IEnumerable<string> GetAuthorNamesAsync(IEnumerable<TeamIdeaEntity> teamIdeas);

        /// <summary>
        /// Get combined query to fetch team ideas as per the selected filter.
        /// </summary>
        /// <param name="postTypes">Post type like: Blog post or Other.</param>
        /// <param name="sharedByNames">User names selected in filter.</param>
        /// <returns>Represents user names query to filter team posts.</returns>
        string GetFilterSearchQuery(string postTypes, string sharedByNames);

        /// <summary>
        /// Get filtered category Ids from team ideas data.
        /// </summary>
        /// <param name="teamIdeas">Represents a collection of team ideas.</param>
        /// <returns>Represents team posts.</returns>
        IEnumerable<string> GetCategoryIds(IEnumerable<TeamIdeaEntity> teamIdeas);

        /// <summary>
        /// Get tags to fetch team posts as per the configured categories.
        /// </summary>
        /// <param name="categories">Categories of a configured team post.</param>
        /// <returns>Represents categories to fetch team ideas.</returns>
        string GetCategories(string categories);

        /// <summary>
        /// Get filtered tag names from team ideas data.
        /// </summary>
        /// <param name="teamIdeas">Represents a collection of team ideas.</param>
        /// <returns>Represents collection of tag names.</returns>
        IEnumerable<string> GetTeamTagsNamesAsync(IEnumerable<TeamIdeaEntity> teamIdeas);

        /// <summary>
        /// Get idea category query to fetch team ideas as per the selected filter.
        /// </summary>
        /// <param name="categories">Team's configured categories.</param>
        /// <returns>Represents post type query to filter team posts.</returns>
        string GetIdeaCategoriesQuery(string categories);
    }
}
