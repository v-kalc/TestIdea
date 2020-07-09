// <copyright file="TeamIdeaStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Common.Providers
{
    using System;
    using System.Linq;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces;
    using Microsoft.Teams.Apps.SubmitIdea.Models;
    using Microsoft.Teams.Apps.SubmitIdea.Models.Configuration;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Implements storage provider which helps to create, get, update or delete team idea data in Microsoft Azure Table storage.
    /// </summary>
    public class TeamIdeaStorageProvider : BaseStorageProvider, ITeamIdeaStorageProvider
    {
        /// <summary>
        /// Represents team idea entity name.
        /// </summary>
        private const string TeamIdeaEntityName = "TeamIdeaEntity";

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamIdeaStorageProvider"/> class.
        /// Handles Microsoft Azure Table storage read write operations.
        /// </summary>
        /// <param name="options">A set of key/value application configuration properties for Microsoft Azure Table storage.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        public TeamIdeaStorageProvider(
            IOptions<StorageSettings> options,
            ILogger<BaseStorageProvider> logger)
            : base(options?.Value.ConnectionString, TeamIdeaEntityName, logger)
        {
            options = options ?? throw new ArgumentNullException(nameof(options));
        }

        /// <summary>
        /// Stores or update team idea details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamIdeaEntity">Holds team idea detail entity data.</param>
        /// <returns>A boolean that represents team idea entity data is successfully saved/updated or not.</returns>
        public async Task<bool> UpsertTeamIdeaAsync(TeamIdeaEntity teamIdeaEntity)
        {
            var result = await this.StoreOrUpdateEntityAsync(teamIdeaEntity);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }

        /// <summary>
        /// Get team idea data from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="ideaId">Idea id to fetch the post details.</param>
        /// <returns>A task that represent a object to hold team post data.</returns>
        public async Task<TeamIdeaEntity> GetTeamIdeaEntityAsync(string ideaId)
        {
            // When there is no team post created by user and Messaging Extension is open, table initialization is required here before creating search index or data source or indexer.
            await this.EnsureInitializedAsync();

            if (string.IsNullOrEmpty(ideaId))
            {
                return null;
            }

            string partitionKeyCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, TeamIdeaEntityName);
            string postIdCondition = TableQuery.GenerateFilterCondition("RowKey", QueryComparisons.Equal, ideaId);
            var combinedPartitionFilter = TableQuery.CombineFilters(partitionKeyCondition, TableOperators.And, postIdCondition);

            TableQuery<TeamIdeaEntity> query = new TableQuery<TeamIdeaEntity>().Where(combinedPartitionFilter);
            var queryResult = await this.SubmitIdeaCloudTable.ExecuteQuerySegmentedAsync(query, null);

            return queryResult?.FirstOrDefault();
        }

        /// <summary>
        /// Get post data.
        /// </summary>
        /// <param name="postCreatedByuserId">User id to fetch the post details.</param>
        /// <param name="postId">Post id to fetch the post details.</param>
        /// <returns>A task that represent a object to hold post data.</returns>
        public async Task<TeamIdeaEntity> GetPostAsync(string postCreatedByuserId, string postId)
        {
            // When there is no post created by user and Messaging Extension is open, table initialization is required here before creating search index or data source or indexer.
            await this.EnsureInitializedAsync();

            if (string.IsNullOrEmpty(postId) || string.IsNullOrEmpty(postCreatedByuserId))
            {
                return null;
            }

            string partitionKeyCondition = TableQuery.GenerateFilterCondition(nameof(TeamIdeaEntity.CreatedByObjectId), QueryComparisons.Equal, postCreatedByuserId);
            string postIdCondition = TableQuery.GenerateFilterCondition("RowKey", QueryComparisons.Equal, postId);
            var combinedPartitionFilter = TableQuery.CombineFilters(partitionKeyCondition, TableOperators.And, postIdCondition);

            TableQuery<TeamIdeaEntity> query = new TableQuery<TeamIdeaEntity>().Where(combinedPartitionFilter);
            var queryResult = await this.SubmitIdeaCloudTable.ExecuteQuerySegmentedAsync(query, null);

            return queryResult?.FirstOrDefault();
        }

        /// <summary>
        /// Stores or update team idea details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="entity">Holds team idea detail entity data.</param>
        /// <returns>A task that represents idea post entity data is saved or updated.</returns>
        private async Task<TableResult> StoreOrUpdateEntityAsync(TeamIdeaEntity entity)
        {
            await this.EnsureInitializedAsync();
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(entity);
            return await this.SubmitIdeaCloudTable.ExecuteAsync(addOrUpdateOperation);
        }
    }
}
