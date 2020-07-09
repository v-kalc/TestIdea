// <copyright file="TeamIdeaController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.SubmitIdea.Common;
    using Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces;
    using Microsoft.Teams.Apps.SubmitIdea.Models;

    /// <summary>
    /// Controller to handle team idea API operations.
    /// </summary>
    [ApiController]
    [Route("api/teamidea")]
    [Authorize]
    public class TeamIdeaController : BaseSubmitIdeaController
    {
        /// <summary>
        /// Event name for team idea HTTP get call.
        /// </summary>
        private const string RecordTeamIdeaHTTPGetCall = "Team idea - HTTP Get call succeeded";

        /// <summary>
        /// Event name for filtered team idea HTTP get call.
        /// </summary>
        private const string RecordFilteredTeamIdeasHTTPGetCall = "Filtered team idea - HTTP Get call succeeded";

        /// <summary>
        /// Event name for team idea unique names HTTP get call.
        /// </summary>
        private const string RecordUniqueUserNamesHTTPGetCall = "Team idea unique user names - HTTP Get call succeeded";

        /// <summary>
        /// Event name for searched team idea for filter HTTP get call.
        /// </summary>
        private const string RecordSearchedTeamIdeasForTitleHTTPGetCall = "Team idea title search - HTTP Get call succeeded";

        /// <summary>
        /// Event name for team idea applied filters HTTP get call.
        /// </summary>
        private const string RecordAppliedFiltersTeamIdeasHTTPGetCall = "Team idea applied filters - HTTP Get call succeeded";

        /// <summary>
        /// Event name for team idea search HTTP get call.
        /// </summary>
        private const string RecordSearchIdeasHTTPGetCall = "Team idea search result - HTTP Get call succeeded";

        /// <summary>
        /// Event name for team idea unique author names HTTP get call.
        /// </summary>
        private const string RecordAuthorNamesHTTPGetCall = "Team post unique author names - HTTP Get call succeeded";

        /// <summary>
        /// Event name for team idea HTTP put call.
        /// </summary>
        private const string RecordTeamIdeaHTTPPatchCall = "Team idea - HTTP Patch call succeeded";

        /// <summary>
        /// Event name for team idea HTTP post call.
        /// </summary>
        private const string RecordTeamIdeaHTTPPostCall = "Team idea - HTTP Post call succeeded";

        /// <summary>
        /// Event name for team idea HTTP post call.
        /// </summary>
        private const string RecordCategoryHTTPGetCall = "Team idea unique category- HTTP get call succeeded";

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Instance of team post storage helper to update post and get information of ideas.
        /// </summary>
        private readonly ITeamIdeaStorageHelper teamIdeaStorageHelper;

        /// <summary>
        /// Instance of team post storage provider to update post and get information of ideas.
        /// </summary>
        private readonly ITeamIdeaStorageProvider teamIdeaStorageProvider;

        /// <summary>
        /// Instance of Search service for working with Microsoft Azure Table storage.
        /// </summary>
        private readonly ITeamIdeaSearchService teamIdeaSearchService;

        /// <summary>
        /// Instance of team tags storage provider for team's discover posts.
        /// </summary>
        private readonly ITeamTagStorageProvider teamTagStorageProvider;

        /// <summary>
        /// Instance of team category storage provider for team's configured categories.
        /// </summary>
        private readonly ITeamCategoryStorageProvider teamCategoryStorageProvider;

        /// <summary>
        /// Instance of team category storage provider for team's configured categories.
        /// </summary>
        private readonly ICategoryStorageProvider categoryStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamIdeaController"/> class.
        /// </summary>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        /// <param name="teamIdeaStorageHelper">Team post storage helper dependency injection.</param>
        /// <param name="teamIdeaStorageProvider">Team idea storage provider dependency injection.</param>
        /// <param name="teamIdeaSearchService">The team post search service dependency injection.</param>
        /// <param name="teamTagStorageProvider">Team tags storage provider dependency injection.</param>
        /// <param name="teamCategoryStorageProvider">Team category storage provider dependency injection.</param>
        /// <param name="categoryStorageProvider">Category storage provider dependency injection.</param>
        public TeamIdeaController(
            ILogger<TeamIdeaController> logger,
            TelemetryClient telemetryClient,
            ITeamIdeaStorageHelper teamIdeaStorageHelper,
            ITeamIdeaStorageProvider teamIdeaStorageProvider,
            ITeamIdeaSearchService teamIdeaSearchService,
            ITeamTagStorageProvider teamTagStorageProvider,
            ITeamCategoryStorageProvider teamCategoryStorageProvider,
            ICategoryStorageProvider categoryStorageProvider)
            : base(telemetryClient)
        {
            this.logger = logger;
            this.teamIdeaStorageHelper = teamIdeaStorageHelper;
            this.teamIdeaSearchService = teamIdeaSearchService;
            this.teamTagStorageProvider = teamTagStorageProvider;
            this.teamIdeaStorageProvider = teamIdeaStorageProvider;
            this.teamCategoryStorageProvider = teamCategoryStorageProvider;
            this.categoryStorageProvider = categoryStorageProvider;
        }

        /// <summary>
        /// Get call to retrieve list of team posts.
        /// </summary>
        /// <param name="pageCount">Page number to get search data from Azure Search service.</param>
        /// <returns>List of team posts.</returns>
        [HttpGet]
        public async Task<IActionResult> GetAsync(int pageCount)
        {
            this.logger.LogInformation("Call to retrieve list of team posts.");
            if (pageCount < 0)
            {
                this.logger.LogError("Invalid value for argument pageCount.");
                return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Invalid value for argument pageCount.");
            }

            var skipRecords = pageCount * Constants.LazyLoadPerPagePostCount;

            try
            {
                var teamPosts = await this.teamIdeaSearchService.GetTeamIdeasAsync(TeamPostSearchScope.AllItems, searchQuery: null, userObjectId: null, count: Constants.LazyLoadPerPagePostCount, skip: skipRecords);
                this.RecordEvent(RecordTeamIdeaHTTPGetCall);

                return this.Ok(teamPosts);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to team post service.");
                throw;
            }
        }

        /// <summary>
        /// Get call to retrieve team idea entity.
        /// </summary>
        /// <param name="ideaId">Unique id of idea to get fetch data from Azure Table storage.</param>
        /// <returns>List of team posts.</returns>
        [HttpGet("idea")]
        public async Task<IActionResult> GetIdeaAsync(string ideaId)
        {
            try
            {
                this.logger.LogInformation("Call to retrieve idea by idea Id.");
                if (string.IsNullOrEmpty(ideaId))
                {
                    this.logger.LogError("Invalid value for argument ideaId.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Invalid value for argument ideaId.");
                }

                var teamPosts = await this.teamIdeaStorageProvider.GetTeamIdeaEntityAsync(ideaId);
                this.RecordEvent(RecordTeamIdeaHTTPGetCall);

                return this.Ok(teamPosts);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to team post service.");
                throw;
            }
        }

        /// <summary>
        /// Post call to store team posts details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamIdeaEntity">Holds team idea detail entity data.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost]
        public async Task<IActionResult> PostAsync([FromBody] TeamIdeaEntity teamIdeaEntity)
        {
            try
            {
                this.logger.LogInformation("Call to add team posts details.");
                teamIdeaEntity = teamIdeaEntity ?? throw new ArgumentNullException(nameof(teamIdeaEntity));

                var updatedTeamPostEntity = new TeamIdeaEntity
                {
                    IdeaId = Guid.NewGuid().ToString(),
                    CreatedByObjectId = this.UserAadId,
                    CreatedByName = this.UserName,
                    CreatedDate = DateTime.UtcNow,
                    UpdatedDate = DateTime.UtcNow,
                    Title = teamIdeaEntity.Title,
                    Description = teamIdeaEntity.Description,
                    Category = teamIdeaEntity.Category,
                    CategoryId = teamIdeaEntity.CategoryId,
                    DocumentLinks = teamIdeaEntity.DocumentLinks,
                    Tags = teamIdeaEntity.Tags,
                    TotalVotes = 0,
                    Status = 0,
                    CreatedByUserPrincipleName = teamIdeaEntity.CreatedByUserPrincipleName,
                };

                var result = await this.teamIdeaStorageProvider.UpsertTeamIdeaAsync(updatedTeamPostEntity);

                if (result)
                {
                    this.RecordEvent(RecordTeamIdeaHTTPPostCall);
                    await this.teamIdeaSearchService.RunIndexerOnDemandAsync();

                    return this.Ok(updatedTeamPostEntity);
                }

                return this.Ok(result);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to team post service.");
                throw;
            }
        }

        /// <summary>
        /// Put call to update team post details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamIdeaEntity">Holds team idea detail entity data.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPatch]
        public async Task<IActionResult> PatchAsync([FromBody] TeamIdeaEntity teamIdeaEntity)
        {
            try
            {
                teamIdeaEntity = teamIdeaEntity ?? throw new ArgumentNullException(nameof(teamIdeaEntity));

                this.logger.LogInformation("Call to update team post details.");

                if (string.IsNullOrEmpty(teamIdeaEntity.IdeaId))
                {
                    this.logger.LogError("Error while updating team post details data in Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while updating team post details data");
                }

                // Validating Idea Id as it will be generated at server side in case of adding new post but cannot be null or empty in case of update.
                var currentPost = await this.teamIdeaStorageProvider.GetTeamIdeaEntityAsync(teamIdeaEntity.IdeaId);
                if (currentPost == null)
                {
                    this.logger.LogError($"User {this.UserAadId} is forbidden to update idea {teamIdeaEntity.IdeaId}.");
                    this.RecordEvent("Update idea - HTTP Patch call failed");

                    return this.Forbid($"You do not have required access to update idea {teamIdeaEntity.IdeaId}.");
                }

                teamIdeaEntity.ApprovedOrRejectedByName = this.UserName;
                var result = await this.teamIdeaStorageProvider.UpsertTeamIdeaAsync(teamIdeaEntity);

                if (result)
                {
                    this.RecordEvent(RecordTeamIdeaHTTPPatchCall);
                    await this.teamIdeaSearchService.RunIndexerOnDemandAsync();
                }

                return this.Ok(result);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to team post service.");
                throw;
            }
        }

        /// <summary>
        /// Get unique user names from Microsoft Azure Table storage.
        /// </summary>
        /// <returns>Returns unique user names.</returns>
        [HttpGet("unique-user-names")]
        public async Task<IActionResult> GetUniqueUserNamesAsync()
        {
            try
            {
                this.logger.LogInformation("Call to get unique names.");

                // Search query will be null if there is no search criteria used. userObjectId will be used when we want to get posts created by respective user.
                var teamPosts = await this.teamIdeaSearchService.GetTeamIdeasAsync(TeamPostSearchScope.UniqueUserNames, searchQuery: null, userObjectId: null);
                var authorNames = this.teamIdeaStorageHelper.GetAuthorNamesAsync(teamPosts);

                this.RecordEvent(RecordUniqueUserNamesHTTPGetCall);

                return this.Ok(authorNames);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to get unique user names.");
                throw;
            }
        }

        /// <summary>
        /// Get list of team posts as per the title text.
        /// </summary>
        /// <param name="searchText">Search text represents the title field to find and get team posts.</param>
        /// <param name="pageCount">Page number to get search data from Azure Search service.</param>
        /// <returns>List of filtered team posts as per the search text for title.</returns>
        [HttpGet("search-team-posts")]
        public async Task<IActionResult> GetSearchedTeamPostsForTitleAsync(string searchText, int pageCount)
        {
            if (pageCount < 0)
            {
                this.logger.LogError("Invalid argument value for pageCount.");
                return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Invalid argument value for pageCount.");
            }

            var skipRecords = pageCount * Constants.LazyLoadPerPagePostCount;

            try
            {
                this.logger.LogInformation("Call to get list of team posts.");
                var teamPosts = await this.teamIdeaSearchService.GetTeamIdeasAsync(TeamPostSearchScope.SearchTeamPostsForTitleText, searchText, userObjectId: null, skip: skipRecords, count: Constants.LazyLoadPerPagePostCount);
                this.RecordEvent(RecordSearchedTeamIdeasForTitleHTTPGetCall);

                return this.Ok(teamPosts);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to get team posts for search title text.");
                throw;
            }
        }

        /// <summary>
        /// Get team posts as per the applied filters from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="categories">Semicolon separated types of posts like blog post or Other.</param>
        /// /// <param name="sharedByNames">Semicolon separated User names to filter the posts.</param>
        /// /// <param name="tags">Semicolon separated tags to match the post tags for which data will fetch.</param>
        /// /// <param name="sortBy">Represents sorting type like: Popularity or Newest.</param>
        /// <param name="teamId">Team id to get configured tags for a team.</param>
        /// <param name="pageCount">Page count for which post needs to be fetched.</param>
        /// <returns>Returns filtered list of team posts as per the selected filters.</returns>
        [HttpGet("applied-filtered-team-posts")]
        public async Task<IActionResult> GetAppliedFiltersTeamIdeasAsync(string categories, string sharedByNames, string tags, string sortBy, string teamId, int pageCount)
        {
            if (pageCount < 0)
            {
                this.logger.LogError("Invalid argument value for pageCount.");
                return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Invalid argument value for pageCount.");
            }

            var skipRecords = pageCount * Constants.LazyLoadPerPagePostCount;

            var teamCategoryEntity = new TeamCategoryEntity();

            try
            {
                this.logger.LogInformation("Call to get team posts as per the applied filters.");

                // Team id will be empty when called from personal scope Discover tab.
                if (!string.IsNullOrEmpty(teamId))
                {
                    teamCategoryEntity = await this.teamCategoryStorageProvider.GetTeamCategoriesDataAsync(teamId);
                }

                var tagsQuery = string.IsNullOrEmpty(tags) ? "*" : this.teamIdeaStorageHelper.GetTags(tags);
                categories = string.IsNullOrEmpty(categories) ? teamCategoryEntity.Categories : categories;
                var filterQuery = this.teamIdeaStorageHelper.GetFilterSearchQuery(categories, sharedByNames);
                var teamPosts = await this.teamIdeaSearchService.GetTeamIdeasAsync(TeamPostSearchScope.FilterTeamPosts, tagsQuery, userObjectId: null, sortBy: sortBy, filterQuery: filterQuery, count: Constants.LazyLoadPerPagePostCount, skip: skipRecords);

                this.RecordEvent(RecordAppliedFiltersTeamIdeasHTTPGetCall);

                return this.Ok(teamPosts);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to get team posts as per the applied filters service.");
                throw;
            }
        }

        /// <summary>
        /// Get list of posts for team's discover tab, as per the configured tags and title of posts.
        /// </summary>
        /// <param name="searchText">Search text represents the title of the posts.</param>
        /// <param name="teamId">Team id to get configured tags for a team.</param>
        /// <param name="pageCount">Page count for which post needs to be fetched.</param>
        /// <returns>List of posts as per the title and configured tags.</returns>
        [HttpGet("team-search-posts")]
        public async Task<IActionResult> GetTeamDiscoverSearchPostsAsync(string searchText, string teamId, int pageCount)
        {
            var skipRecords = pageCount * Constants.LazyLoadPerPagePostCount;

            try
            {
                this.logger.LogInformation("Call to get list of posts as per the configured tags and title.");

                if (string.IsNullOrEmpty(teamId))
                {
                    this.logger.LogError("Error while fetching search posts as per the title and configured tags from Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while fetching search posts as per the title and configured tags from Microsoft Azure Table storage.");
                }

                // Get tags based on the team id for which tags has configured.
                var teamCategoryEntity = await this.teamCategoryStorageProvider.GetTeamCategoriesDataAsync(teamId);

                if (teamCategoryEntity == null || string.IsNullOrEmpty(teamCategoryEntity.Categories))
                {
                    return this.Ok();
                }

                var categoriesQuery = string.IsNullOrEmpty(teamCategoryEntity.Categories) ? "*" : this.teamIdeaStorageHelper.GetCategories(teamCategoryEntity.Categories);
                var filterQuery = $"search.ismatch('{categoriesQuery}', 'CategoryId')";
                var teamPosts = await this.teamIdeaSearchService.GetTeamIdeasAsync(TeamPostSearchScope.SearchTeamPostsForTitleText, searchText, userObjectId: null, count: Constants.LazyLoadPerPagePostCount, skip: skipRecords, filterQuery: filterQuery);
                this.RecordEvent(RecordSearchIdeasHTTPGetCall);

                return this.Ok(teamPosts);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to get team posts for search title text.");
                throw;
            }
        }

        /// <summary>
        /// Get unique author names from storage.
        /// </summary>
        /// <param name="teamId">Team id to get the configured categories for a team.</param>
        /// <returns>Returns unique user names.</returns>
        [HttpGet("authors-for-categories")]
        public async Task<IActionResult> GetAuthorNamesAsync(string teamId)
        {
            try
            {
                this.logger.LogInformation("Call to get unique author names.");

                var names = new List<string>();

                // Get tags based on the team id for which tags has configured.
                var teamCategoryEntity = await this.teamCategoryStorageProvider.GetTeamCategoriesDataAsync(teamId);

                if (teamCategoryEntity == null || string.IsNullOrEmpty(teamCategoryEntity.Categories))
                {
                    return this.Ok(names);
                }

                var tagsQuery = this.teamIdeaStorageHelper.GetCategories(teamCategoryEntity.Categories);
                var teamIdeas = await this.teamIdeaSearchService.GetTeamIdeasAsync(TeamPostSearchScope.FilterAsPerTeamTags, tagsQuery, null, null);
                var authorNames = this.teamIdeaStorageHelper.GetAuthorNamesAsync(teamIdeas);
                this.RecordEvent(RecordAuthorNamesHTTPGetCall);

                return this.Ok(authorNames);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to get unique user names.");
                throw;
            }
        }

        /// <summary>
        /// Get list of unique tags to show while configuring the preference.
        /// </summary>
        /// <param name="searchText">Search text represents the text to find and get unique tags.</param>
        /// <returns>List of unique tags.</returns>
        [HttpGet("unique-categories")]
        public async Task<IActionResult> GetUniqueCategoriesAsync(string searchText)
        {
            try
            {
                this.logger.LogInformation("Call to get list of unique tags to show while configuring the preference.");

                if (string.IsNullOrEmpty(searchText))
                {
                    this.logger.LogError("Error while getting the list of unique tags from Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while getting the list of unique tags from Microsoft Azure Table storage.");
                }

                var teamPosts = await this.teamIdeaSearchService.GetTeamIdeasAsync(TeamPostSearchScope.Categories, searchText, userObjectId: null);
                var uniqueCategoryIds = this.teamIdeaStorageHelper.GetCategoryIds(teamPosts);
                var categories = await this.categoryStorageProvider.GetCategoriesByIdsAsync(uniqueCategoryIds);
                this.RecordEvent(RecordCategoryHTTPGetCall);

                return this.Ok(categories);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to get unique tags.");
                throw;
            }
        }

        /// <summary>
        /// Get unique tags for teams categories from storage.
        /// </summary>
        /// <param name="teamId">Team id to get the configured categories for a team.</param>
        /// <returns>Returns unique tags for given team categories.</returns>
        [HttpGet("tags-for-categories")]
        public async Task<IActionResult> GetTeamCategoryTagsAsync(string teamId)
        {
            try
            {
                this.logger.LogInformation("Call to get unique author names.");

                var names = new List<string>();

                // Get tags based on the team id for which tags has configured.
                var teamCategoryEntity = await this.teamCategoryStorageProvider.GetTeamCategoriesDataAsync(teamId);

                if (teamCategoryEntity == null || string.IsNullOrEmpty(teamCategoryEntity.Categories))
                {
                    return this.Ok(names);
                }

                var tagsQuery = this.teamIdeaStorageHelper.GetCategories(teamCategoryEntity.Categories);
                var teamIdeas = await this.teamIdeaSearchService.GetTeamIdeasAsync(TeamPostSearchScope.FilterAsPerTeamTags, tagsQuery, null, null);
                var authorNames = this.teamIdeaStorageHelper.GetTeamTagsNamesAsync(teamIdeas);
                this.RecordEvent(RecordAuthorNamesHTTPGetCall);

                return this.Ok(authorNames);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to get unique user names.");
                throw;
            }
        }

        /// <summary>
        /// Get list of teams configured unique categories.
        /// </summary>
        /// <param name="teamId">team identifier.</param>
        /// <returns>List of unique categories.</returns>
        [HttpGet("team-unique-categories")]
        public async Task<IActionResult> GetTeamsUniqueCategoriesAsync(string teamId)
        {
            try
            {
                this.logger.LogInformation("Call to get list of unique tags to show while configuring the preference.");

                if (string.IsNullOrEmpty(teamId))
                {
                    this.logger.LogError("Error while getting the list of unique tags from Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while getting the list of unique tags from Microsoft Azure Table storage.");
                }

                var teamCategoryEntity = await this.teamCategoryStorageProvider.GetTeamCategoriesDataAsync(teamId);
                var categoryIds = teamCategoryEntity.Categories.Split(';').Where(categoryId => !string.IsNullOrWhiteSpace(categoryId)).Select(categoryId => categoryId.Trim());
                var categories = await this.categoryStorageProvider.GetCategoriesByIdsAsync(categoryIds);
                this.RecordEvent(RecordCategoryHTTPGetCall);

                return this.Ok(categories);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to get unique tags.");
                throw;
            }
        }

        /// <summary>
        /// Get filtered team ideas for particular team as per the configured categories.
        /// </summary>
        /// <param name="teamId">Team id for which data will fetch.</param>
        /// <param name="pageCount">Page number to get search data.</param>
        /// <returns>Returns filtered list of team posts as per the configured tags.</returns>
        [HttpGet("team-ideas")]
        public async Task<IActionResult> GetFilteredTeamPostsAsync(string teamId, int pageCount)
        {
            var skipRecords = pageCount * Constants.LazyLoadPerPagePostCount;

            try
            {
                this.logger.LogInformation("Call to get filtered team idea details.");

                if (string.IsNullOrEmpty(teamId))
                {
                    this.logger.LogError("Error while fetching filtered team posts as per the configured tags.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while fetching filtered team ideas as per the configured categories.");
                }

                IEnumerable<TeamIdeaEntity> teamIdeas = new List<TeamIdeaEntity>();

                // Get categories based on the team id for which categories has configured.
                var teamCategories = await this.teamCategoryStorageProvider.GetTeamCategoriesDataAsync(teamId);

                if (teamCategories == null || string.IsNullOrEmpty(teamCategories.Categories))
                {
                    return this.Ok(teamIdeas);
                }

                // Prepare query based on the tags and get the data using search service.
                var categoriesQuery = this.teamIdeaStorageHelper.GetCategories(teamCategories.Categories);
                teamIdeas = await this.teamIdeaSearchService.GetTeamIdeasAsync(TeamPostSearchScope.FilterAsPerTeamTags, categoriesQuery, userObjectId: null, count: Constants.LazyLoadPerPagePostCount, skip: skipRecords);

                this.RecordEvent(RecordFilteredTeamIdeasHTTPGetCall);

                return this.Ok(teamIdeas);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to team post service.");
                throw;
            }
        }
    }
}