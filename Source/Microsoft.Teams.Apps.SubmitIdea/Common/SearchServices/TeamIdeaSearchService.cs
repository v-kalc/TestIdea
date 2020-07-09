// <copyright file="TeamIdeaSearchService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Common.SearchServices
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Azure.Search;
    using Microsoft.Azure.Search.Models;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces;
    using Microsoft.Teams.Apps.SubmitIdea.Models;
    using Microsoft.Teams.Apps.SubmitIdea.Models.Configuration;

    /// <summary>
    /// Team idea Search service which helps in creating index, indexer and data source if it doesn't exist
    /// for indexing table which will be used for search by Messaging Extension.
    /// </summary>
    public class TeamIdeaSearchService : ITeamIdeaSearchService, IDisposable
    {
        /// <summary>
        /// Azure Search service index name for team post.
        /// </summary>
        private const string TeamIdeaIndexName = "team-idea-index";

        /// <summary>
        /// Azure Search service indexer name for team post.
        /// </summary>
        private const string TeamIdeaIndexerName = "team-idea-indexer";

        /// <summary>
        /// Azure Search service data source name for team post.
        /// </summary>
        private const string TeamIdeaDataSourceName = "team-idea-storage";

        /// <summary>
        /// Table name where team post data will get saved.
        /// </summary>
        private const string TeamIdeaTableName = "TeamIdeaEntity";

        /// <summary>
        /// Represents the sorting type as popularity means to sort the data based on number of votes.
        /// </summary>
        private const string SortByPopular = "Popularity";

        /// <summary>
        /// Azure Search service maximum search result count for team post entity.
        /// </summary>
        private const int ApiSearchResultCount = 1500;

        /// <summary>
        /// Used to initialize task.
        /// </summary>
        private readonly Lazy<Task> initializeTask;

        /// <summary>
        /// Instance of Azure Search service client.
        /// </summary>
        private readonly SearchServiceClient searchServiceClient;

        /// <summary>
        /// Instance of Azure Search index client.
        /// </summary>
        private readonly SearchIndexClient searchIndexClient;

        /// <summary>
        /// Instance of team post storage helper to update post and get information of posts.
        /// </summary>
        private readonly ITeamIdeaStorageProvider teamIdeaStorageProvider;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<TeamIdeaSearchService> logger;

        /// <summary>
        /// Represents a set of key/value application configuration properties.
        /// </summary>
        private readonly SearchServiceSettings options;

        /// <summary>
        /// Flag: Has Dispose already been called?
        /// </summary>
        private bool disposed = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamIdeaSearchService"/> class.
        /// </summary>
        /// <param name="optionsAccessor">A set of key/value application configuration properties.</param>
        /// <param name="teamIdeaStorageProvider">Team idea storage provider dependency injection.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="searchServiceClient">Search service client dependency injection.</param>
        /// <param name="searchIndexClient">Search index client dependency injection.</param>
        public TeamIdeaSearchService(
            IOptions<SearchServiceSettings> optionsAccessor,
            ITeamIdeaStorageProvider teamIdeaStorageProvider,
            ILogger<TeamIdeaSearchService> logger,
            SearchServiceClient searchServiceClient,
            SearchIndexClient searchIndexClient)
        {
            optionsAccessor = optionsAccessor ?? throw new ArgumentNullException(nameof(optionsAccessor));
            this.options = optionsAccessor.Value;
            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync());
            this.teamIdeaStorageProvider = teamIdeaStorageProvider;
            this.logger = logger;
            this.searchServiceClient = searchServiceClient;
            this.searchIndexClient = searchIndexClient;
        }

        /// <summary>
        /// Provide search result for table to be used by user's based on Azure Search service.
        /// </summary>
        /// <param name="searchScope">Scope of the search.</param>
        /// <param name="searchQuery">Query which the user had typed in Messaging Extension search field.</param>
        /// <param name="userObjectId">Azure Active Directory object id of the user.</param>
        /// <param name="count">Number of search results to return.</param>
        /// <param name="skip">Number of search results to skip.</param>
        /// <param name="sortBy">Represents sorting type like: Popularity or Newest.</param>
        /// <param name="filterQuery">Filter bar based query.</param>
        /// <returns>List of search results.</returns>
        public async Task<IEnumerable<TeamIdeaEntity>> GetTeamIdeasAsync(
            TeamPostSearchScope searchScope,
            string searchQuery,
            string userObjectId,
            int? count = null,
            int? skip = null,
            string sortBy = null,
            string filterQuery = null)
        {
            try
            {
                await this.EnsureInitializedAsync();
                IEnumerable<TeamIdeaEntity> teamPosts = new List<TeamIdeaEntity>();
                var searchParameters = this.InitializeSearchParameters(searchScope, userObjectId, count, skip, sortBy, filterQuery);

                SearchContinuationToken continuationToken = null;
                var userIdeasCollection = new List<TeamIdeaEntity>();
                var teamPostResult = await this.searchIndexClient.Documents.SearchAsync<TeamIdeaEntity>(searchQuery, searchParameters);

                if (teamPostResult?.Results != null)
                {
                    userIdeasCollection.AddRange(teamPostResult.Results.Select(p => p.Document));
                    continuationToken = teamPostResult.ContinuationToken;
                }

                if (continuationToken == null)
                {
                    return userIdeasCollection;
                }

                do
                {
                    var teamPostResult1 = await this.searchIndexClient.Documents.ContinueSearchAsync<TeamIdeaEntity>(continuationToken);

                    if (teamPostResult1?.Results != null)
                    {
                        userIdeasCollection.AddRange(teamPostResult1.Results.Select(p => p.Document));
                        continuationToken = teamPostResult1.ContinuationToken;
                    }
                }
                while (continuationToken != null);

                return userIdeasCollection;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Creates Index, Data Source and Indexer for search service.
        /// </summary>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task RecreateSearchServiceIndexAsync()
        {
            try
            {
                await this.CreateSearchIndexAsync();
                await this.CreateDataSourceAsync();
                await this.CreateIndexerAsync();
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Run the indexer on demand.
        /// </summary>
        /// <returns>A task that represents the work queued to execute</returns>
        public async Task RunIndexerOnDemandAsync()
        {
            await this.searchServiceClient.Indexers.RunAsync(TeamIdeaIndexerName);
        }

        /// <summary>
        /// Dispose search service instance.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Protected implementation of Dispose pattern.
        /// </summary>
        /// <param name="disposing">True if already disposed else false.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (this.disposed)
            {
                return;
            }

            if (disposing)
            {
                this.searchServiceClient.Dispose();
                this.searchIndexClient.Dispose();
            }

            this.disposed = true;
        }

        /// <summary>
        /// Create index, indexer and data source if doesn't exist.
        /// </summary>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task InitializeAsync()
        {
            try
            {
                // When there is no team post created by user and Messaging Extension is open, table initialization is required here before creating search index or data source or indexer.
                await this.teamIdeaStorageProvider.GetTeamIdeaEntityAsync(string.Empty);
                await this.RecreateSearchServiceIndexAsync();
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Failed to initialize Azure Search Service: {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Create index in Azure Search service if it doesn't exist.
        /// </summary>
        /// <returns><see cref="Task"/> That represents index is created if it is not created.</returns>
        private async Task CreateSearchIndexAsync()
        {
            if (await this.searchServiceClient.Indexes.ExistsAsync(TeamIdeaIndexName))
            {
                await this.searchServiceClient.Indexes.DeleteAsync(TeamIdeaIndexName);
            }

            var tableIndex = new Index()
            {
                Name = TeamIdeaIndexName,
                Fields = FieldBuilder.BuildForType<TeamIdeaEntity>(),
            };
            await this.searchServiceClient.Indexes.CreateAsync(tableIndex);
        }

        /// <summary>
        /// Create data source if it doesn't exist in Azure Search service.
        /// </summary>
        /// <returns><see cref="Task"/> That represents data source is added to Azure Search service.</returns>
        private async Task CreateDataSourceAsync()
        {
            if (await this.searchServiceClient.DataSources.ExistsAsync(TeamIdeaDataSourceName))
            {
                return;
            }

            var dataSource = DataSource.AzureTableStorage(
                TeamIdeaDataSourceName,
                this.options.ConnectionString,
                TeamIdeaTableName);

            await this.searchServiceClient.DataSources.CreateAsync(dataSource);
        }

        /// <summary>
        /// Create indexer if it doesn't exist in Azure Search service.
        /// </summary>
        /// <returns><see cref="Task"/> That represents indexer is created if not available in Azure Search service.</returns>
        private async Task CreateIndexerAsync()
        {
            if (await this.searchServiceClient.Indexers.ExistsAsync(TeamIdeaIndexerName))
            {
                await this.searchServiceClient.Indexers.DeleteAsync(TeamIdeaIndexerName);
            }

            var indexer = new Indexer()
            {
                Name = TeamIdeaIndexerName,
                DataSourceName = TeamIdeaDataSourceName,
                TargetIndexName = TeamIdeaIndexName,
            };

            await this.searchServiceClient.Indexers.CreateAsync(indexer);
            await this.searchServiceClient.Indexers.RunAsync(TeamIdeaIndexerName);
        }

        /// <summary>
        /// Initialization of InitializeAsync method which will help in indexing.
        /// </summary>
        /// <returns>Represents an asynchronous operation.</returns>
        private Task EnsureInitializedAsync()
        {
            return this.initializeTask.Value;
        }

        /// <summary>
        /// Initialization of search service parameters which will help in searching the documents.
        /// </summary>
        /// <param name="searchScope">Scope of the search.</param>
        /// <param name="userObjectId">Azure Active Directory object id of the user.</param>
        /// <param name="count">Number of search results to return.</param>
        /// <param name="skip">Number of search results to skip.</param>
        /// <param name="sortBy">Represents sorting type like: Popularity or Newest.</param>
        /// <param name="filterQuery">Filter bar based query.</param>
        /// <returns>Represents an search parameter object.</returns>
        private SearchParameters InitializeSearchParameters(
            TeamPostSearchScope searchScope,
            string userObjectId,
            int? count = null,
            int? skip = null,
            string sortBy = null,
            string filterQuery = null)
        {
            SearchParameters searchParameters = new SearchParameters()
            {
                Top = count ?? ApiSearchResultCount,
                Skip = skip ?? 0,
                IncludeTotalResultCount = false,
                Select = new[]
                {
                    nameof(TeamIdeaEntity.IdeaId),
                    nameof(TeamIdeaEntity.CategoryId),
                    nameof(TeamIdeaEntity.Category),
                    nameof(TeamIdeaEntity.Title),
                    nameof(TeamIdeaEntity.Description),
                    nameof(TeamIdeaEntity.Tags),
                    nameof(TeamIdeaEntity.CreatedDate),
                    nameof(TeamIdeaEntity.CreatedByName),
                    nameof(TeamIdeaEntity.UpdatedDate),
                    nameof(TeamIdeaEntity.CreatedByObjectId),
                    nameof(TeamIdeaEntity.TotalVotes),
                    nameof(TeamIdeaEntity.Status),
                    nameof(TeamIdeaEntity.CreatedByUserPrincipleName),
                },
                SearchFields = new[] { nameof(TeamIdeaEntity.Title) },
                Filter = string.IsNullOrEmpty(filterQuery) ? null : $"({filterQuery})",
            };

            switch (searchScope)
            {
                case TeamPostSearchScope.AllItems:
                    searchParameters.OrderBy = new[] { $"{nameof(TeamIdeaEntity.UpdatedDate)} desc" };

                    break;

                case TeamPostSearchScope.PostedByMe:
                    searchParameters.Filter = $"{nameof(TeamIdeaEntity.CreatedByObjectId)} eq '{userObjectId}' ";
                    searchParameters.OrderBy = new[] { $"{nameof(TeamIdeaEntity.UpdatedDate)} desc" };
                    break;

                case TeamPostSearchScope.Popular:
                    searchParameters.OrderBy = new[] { $"{nameof(TeamIdeaEntity.TotalVotes)} desc" };
                    break;

                case TeamPostSearchScope.TeamPreferenceTags:
                    searchParameters.SearchFields = new[] { nameof(TeamIdeaEntity.Tags) };
                    searchParameters.Top = 5000;
                    searchParameters.Select = new[] { nameof(TeamIdeaEntity.Tags) };
                    break;

                case TeamPostSearchScope.Categories:
                    searchParameters.SearchFields = new[] { nameof(TeamIdeaEntity.CategoryId) };
                    searchParameters.Top = 5000;
                    searchParameters.Select = new[] { nameof(TeamIdeaEntity.CategoryId) };
                    break;

                case TeamPostSearchScope.FilterAsPerTeamTags:
                    searchParameters.OrderBy = new[] { $"{nameof(TeamIdeaEntity.UpdatedDate)} desc" };
                    searchParameters.SearchFields = new[] { nameof(TeamIdeaEntity.CategoryId) };
                    break;

                case TeamPostSearchScope.FilterPostsAsPerDateRange:
                    searchParameters.OrderBy = new[] { $"{nameof(TeamIdeaEntity.UpdatedDate)} desc" };
                    searchParameters.Top = 200;
                    break;

                case TeamPostSearchScope.UniqueUserNames:
                    searchParameters.OrderBy = new[] { $"{nameof(TeamIdeaEntity.UpdatedDate)} desc" };
                    searchParameters.Select = new[] { nameof(TeamIdeaEntity.CreatedByName) };
                    break;

                case TeamPostSearchScope.SearchTeamPostsForTitleText:
                    searchParameters.OrderBy = new[] { $"{nameof(TeamIdeaEntity.UpdatedDate)} desc" };
                    searchParameters.QueryType = QueryType.Full;
                    searchParameters.SearchFields = new[] { nameof(TeamIdeaEntity.Title) };
                    break;

                case TeamPostSearchScope.Pending:
                    searchParameters.Filter = $"{nameof(TeamIdeaEntity.Status)} eq 0 ";
                    searchParameters.OrderBy = new[] { $"{nameof(TeamIdeaEntity.UpdatedDate)} desc" };
                    break;

                case TeamPostSearchScope.Approved:
                    searchParameters.Filter = $"{nameof(TeamIdeaEntity.Status)} eq 1 ";
                    searchParameters.OrderBy = new[] { $"{nameof(TeamIdeaEntity.UpdatedDate)} desc" };
                    break;

                case TeamPostSearchScope.FilterTeamPosts:

                    if (!string.IsNullOrEmpty(sortBy))
                    {
                        searchParameters.OrderBy = sortBy == SortByPopular ? new[] { $"{nameof(TeamIdeaEntity.TotalVotes)} desc" } : new[] { $"{nameof(TeamIdeaEntity.UpdatedDate)} desc" };
                    }

                    searchParameters.SearchFields = new[] { nameof(TeamIdeaEntity.Tags) };
                    break;
            }

            return searchParameters;
        }
    }
}
