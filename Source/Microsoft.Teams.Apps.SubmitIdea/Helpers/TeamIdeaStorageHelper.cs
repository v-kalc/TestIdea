// <copyright file="TeamIdeaStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces;
    using Microsoft.Teams.Apps.SubmitIdea.Models;

    /// <summary>
    /// Implements team post storage helper which helps to construct the model, create search query for team post.
    /// </summary>
    public class TeamIdeaStorageHelper : ITeamIdeaStorageHelper
    {
        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<TeamIdeaStorageHelper> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamIdeaStorageHelper"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public TeamIdeaStorageHelper(
            ILogger<TeamIdeaStorageHelper> logger)
        {
            this.logger = logger;
        }

        /// <summary>
        /// Create team idea model data.
        /// </summary>
        /// <param name="teamIdeaEntity">Team idea detail.</param>
        /// <param name="userId">User Azure active directory id.</param>
        /// <param name="userName">Author who created the idea.</param>
        /// <returns>A task that represents team idea entity data.</returns>
        public TeamIdeaEntity CreateTeamIdeaModel(TeamIdeaEntity teamIdeaEntity, string userId, string userName)
        {
            try
            {
                teamIdeaEntity = teamIdeaEntity ?? throw new ArgumentNullException(nameof(teamIdeaEntity));

                teamIdeaEntity.IdeaId = Guid.NewGuid().ToString();
                teamIdeaEntity.CreatedByObjectId = userId;
                teamIdeaEntity.CreatedByName = userName;
                teamIdeaEntity.CreatedDate = DateTime.UtcNow;
                teamIdeaEntity.UpdatedDate = DateTime.UtcNow;

                return teamIdeaEntity;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception occurred while creating the team post model data.");
                throw;
            }
        }

        /// <summary>
        /// Create updated team idea model data for Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamIdeaEntity">Team post detail.</param>
        /// <returns>A task that represents team post entity updated data.</returns>
        public TeamIdeaEntity CreateUpdatedTeamIdeaModel(TeamIdeaEntity teamIdeaEntity)
        {
            try
            {
                teamIdeaEntity = teamIdeaEntity ?? throw new ArgumentNullException(nameof(teamIdeaEntity));

                teamIdeaEntity.UpdatedDate = DateTime.UtcNow;

                return teamIdeaEntity;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception occurred while getting the team idea model data");
                throw;
            }
        }

        /// <summary>
        /// Get filtered team posts as per the configured tags.
        /// </summary>
        /// <param name="teamIdeas">Team post entities.</param>
        /// <param name="searchText">Search text for tags.</param>
        /// <returns>Represents team ideas.</returns>
        public IEnumerable<TeamIdeaEntity> GetFilteredTeamIdeasAsPerTags(IEnumerable<TeamIdeaEntity> teamIdeas, string searchText)
        {
            try
            {
                teamIdeas = teamIdeas ?? throw new ArgumentNullException(nameof(teamIdeas));
                searchText = searchText ?? throw new ArgumentNullException(nameof(searchText));
                var filteredTeamIdeas = new List<TeamIdeaEntity>();

                foreach (var teamIdea in teamIdeas)
                {
                    foreach (var tag in searchText.Split(";"))
                    {
                        if (Array.Exists(teamIdea.Tags?.Split(";"), tagText => tagText.Equals(tag.Trim(), StringComparison.InvariantCultureIgnoreCase)))
                        {
                            filteredTeamIdeas.Add(teamIdea);
                            break;
                        }
                    }
                }

                return filteredTeamIdeas;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception occurred while preparing the team preference entities list.");
                throw;
            }
        }

        /// <summary>
        /// Get tags to fetch team posts as per the configured tags.
        /// </summary>
        /// <param name="tags">Tags of a configured team post.</param>
        /// <returns>Represents tags to fetch team posts.</returns>
        public string GetTags(string tags)
        {
            try
            {
                tags = tags ?? throw new ArgumentNullException(nameof(tags));
                var postTags = tags.Split(';').Where(postType => !string.IsNullOrWhiteSpace(postType)).ToList();

                return string.Join(" ", postTags);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception occurred while preparing the query for tags to get team posts as per the configured tags.");
                throw;
            }
        }

        /// <summary>
        /// Get tags to fetch team posts as per the configured categories.
        /// </summary>
        /// <param name="categories">Categories of a configured team post.</param>
        /// <returns>Represents categories to fetch team ideas.</returns>
        public string GetCategories(string categories)
        {
            try
            {
                categories = categories ?? throw new ArgumentNullException(nameof(categories));
                var ideaCategories = categories.Split(';').Where(category => !string.IsNullOrWhiteSpace(category)).Distinct().ToList();

                return string.Join(" ", ideaCategories);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception occurred while preparing the query for categories to get team ideas as per the configured categories.");
                throw;
            }
        }

        /// <summary>
        /// Get filtered team posts as per the date range from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamIdeas">Team ideas data.</param>
        /// <param name="fromDate">Start date from which data should fetch.</param>
        /// <param name="toDate">End date till when data should fetch.</param>
        /// <returns>A task that represent collection to hold team posts data.</returns>
        public IEnumerable<TeamIdeaEntity> GetTeamIdeasInDateRangeAsync(IEnumerable<TeamIdeaEntity> teamIdeas, DateTime fromDate, DateTime toDate)
        {
            return teamIdeas.Where(post => post.UpdatedDate >= fromDate && post.UpdatedDate <= toDate);
        }

        /// <summary>
        /// Get filtered user names from team ideas data.
        /// </summary>
        /// <param name="teamIdeas">Represents a collection of team ideas.</param>
        /// <returns>Represents team posts.</returns>
        public IEnumerable<string> GetAuthorNamesAsync(IEnumerable<TeamIdeaEntity> teamIdeas)
        {
            try
            {
                teamIdeas = teamIdeas ?? throw new ArgumentNullException(nameof(teamIdeas));

                return teamIdeas.Select(idea => idea.CreatedByName).Distinct().OrderBy(createdByName => createdByName);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception occurred while preparing the unique user names list.");
                throw;
            }
        }

        /// <summary>
        /// Get filtered tag names from team ideas data.
        /// </summary>
        /// <param name="teamIdeas">Represents a collection of team ideas.</param>
        /// <returns>Represents collection of tag names.</returns>
        public IEnumerable<string> GetTeamTagsNamesAsync(IEnumerable<TeamIdeaEntity> teamIdeas)
        {
            try
            {
                teamIdeas = teamIdeas ?? throw new ArgumentNullException(nameof(teamIdeas));

                var tagsCollection = new List<string>();
                foreach (var teamIdea in teamIdeas)
                {
                    var tagsData = teamIdea.Tags.Split(';').Distinct();
                    tagsCollection.AddRange(tagsData);
                }

                return tagsCollection.Distinct();
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception occurred while preparing the unique user names list.");
                throw;
            }
        }

        /// <summary>
        /// Get filtered category ids from team ideas data.
        /// </summary>
        /// <param name="teamIdeas">Represents a collection of team ideas.</param>
        /// <returns>Represents team posts.</returns>
        public IEnumerable<string> GetCategoryIds(IEnumerable<TeamIdeaEntity> teamIdeas)
        {
            try
            {
                teamIdeas = teamIdeas ?? throw new ArgumentNullException(nameof(teamIdeas));

                return teamIdeas.Select(idea => idea.CategoryId).Distinct().OrderBy(category => category);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception occurred while preparing the unique user names list.");
                throw;
            }
        }

        /// <summary>
        /// Get combined query to fetch team posts as per the selected filter.
        /// </summary>
        /// <param name="postTypes">Post type like: Blog post or Other.</param>
        /// <param name="sharedByNames">User names selected in filter.</param>
        /// <returns>Represents user names query to filter team posts.</returns>
        public string GetFilterSearchQuery(string postTypes, string sharedByNames)
        {
            try
            {
                var typesQuery = this.GetIdeaCategoriesQuery(postTypes);
                var sharedByNamesQuery = this.GetSharedByNamesQuery(sharedByNames);
                string combinedQuery = string.Empty;

                if (string.IsNullOrEmpty(typesQuery) && string.IsNullOrEmpty(sharedByNamesQuery))
                {
                    return null;
                }

                if (!string.IsNullOrEmpty(typesQuery) && !string.IsNullOrEmpty(sharedByNamesQuery))
                {
                    return $"({typesQuery}) and ({sharedByNamesQuery})";
                }

                if (!string.IsNullOrEmpty(typesQuery))
                {
                    return $"({typesQuery})";
                }

                if (!string.IsNullOrEmpty(sharedByNamesQuery))
                {
                    return $"({sharedByNamesQuery})";
                }

                return null;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception occurred while preparing the query to get filter bar search result for team posts.");
                throw;
            }
        }

        /// <summary>
        /// Get idea category query to fetch team ideas as per the selected filter.
        /// </summary>
        /// <param name="categories">Team's configured categories.</param>
        /// <returns>Represents post type query to filter team posts.</returns>
        public string GetIdeaCategoriesQuery(string categories)
        {
            try
            {
                if (string.IsNullOrEmpty(categories))
                {
                    return null;
                }

                StringBuilder categoryQuery = new StringBuilder();
                var categoryData = categories.Split(';').Where(postType => !string.IsNullOrWhiteSpace(postType)).Select(postType => postType.Trim()).ToList();

                if (categoryData.Count > 1)
                {
                    var posts = categoryData.Take(categoryData.Count - 1).ToList();
                    posts.ForEach(postType =>
                    {
                        categoryQuery.Append($"CategoryId eq '{postType}' or ");
                    });

                    categoryQuery.Append($"CategoryId eq '{categoryData.Last()}'");
                }
                else
                {
                    categoryQuery.Append($"CategoryId eq '{categoryData.Last()}'");
                }

                return categoryQuery.ToString();
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception occurred while preparing the query for post types to get team posts as per the selected types.");
                throw;
            }
        }

        /// <summary>
        /// Get user names query to fetch team posts as per the selected filter.
        /// </summary>
        /// <param name="sharedByNames">User names selected in filter.</param>
        /// <returns>Represents user names query to filter team posts.</returns>
        private string GetSharedByNamesQuery(string sharedByNames)
        {
            try
            {
                if (string.IsNullOrEmpty(sharedByNames))
                {
                    return null;
                }

                StringBuilder sharedByNamesQuery = new StringBuilder();
                var sharedByNamesData = sharedByNames.Split(';').Where(name => !string.IsNullOrWhiteSpace(name)).Select(name => name.Trim()).ToList();

                if (sharedByNamesData.Count > 1)
                {
                    var users = sharedByNamesData.Take(sharedByNamesData.Count - 1).ToList();
                    users.ForEach(user =>
                    {
                        sharedByNamesQuery.Append($"CreatedByName eq '{user}' or ");
                    });

                    sharedByNamesQuery.Append($"CreatedByName eq '{sharedByNamesData.Last()}'");
                }
                else
                {
                    sharedByNamesQuery.Append($"CreatedByName eq '{sharedByNamesData.Last()}'");
                }

                return sharedByNamesQuery.ToString();
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception occurred while preparing the query for shared by names to get team ideas as per the selected names.");
                throw;
            }
        }
    }
}