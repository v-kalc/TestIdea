// <copyright file="TeamPreferenceStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Common.Providers
{
    using System;
    using System.Collections.Generic;
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
    /// Implements storage provider which helps to create, get or update team preferences data in Microsoft Azure Table storage.
    /// </summary>
    public class TeamPreferenceStorageProvider : BaseStorageProvider, ITeamPreferenceStorageProvider
    {
        /// <summary>
        /// Represents team preference entity name.
        /// </summary>
        private const string TeamPreferenceEntityName = "TeamPreferenceEntity";

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamPreferenceStorageProvider"/> class.
        /// Handles Microsoft Azure Table storage read write operations.
        /// </summary>
        /// <param name="options">A set of key/value application configuration properties for Microsoft Azure Table storage.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        public TeamPreferenceStorageProvider(
            IOptions<StorageSettings> options,
            ILogger<BaseStorageProvider> logger)
            : base(options?.Value.ConnectionString, TeamPreferenceEntityName, logger)
        {
            options = options ?? throw new ArgumentNullException(nameof(options));
        }

        /// <summary>
        /// Get team preference data from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamId">Team Id for which need to fetch data.</param>
        /// <returns>A task that represents an object to hold team preference data.</returns>
        public async Task<TeamPreferenceEntity> GetTeamPreferenceDataAsync(string teamId)
        {
            teamId = teamId ?? throw new ArgumentNullException(nameof(teamId));
            await this.EnsureInitializedAsync();

            string partitionKeyCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, TeamPreferenceEntityName);
            string teamIdCondition = TableQuery.GenerateFilterCondition("RowKey", QueryComparisons.Equal, teamId);
            var combinedTeamFilter = TableQuery.CombineFilters(partitionKeyCondition, TableOperators.And, teamIdCondition);

            TableQuery<TeamPreferenceEntity> query = new TableQuery<TeamPreferenceEntity>().Where(combinedTeamFilter);
            var queryResult = await this.SubmitIdeaCloudTable.ExecuteQuerySegmentedAsync(query, null);

            return queryResult?.Results.FirstOrDefault();
        }

        /// <summary>
        /// Get team preferences data from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="digestFrequency">Digest frequency text for notification like Monthly/Weekly.</param>
        /// <returns>A task that represent collection to hold team preferences data.</returns>
        public async Task<IEnumerable<TeamPreferenceEntity>> GetTeamPreferencesAsync(string digestFrequency)
        {
            await this.EnsureInitializedAsync();

            var partitionKeyCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, TeamPreferenceEntityName);
            var digestFrequencyCondition = TableQuery.GenerateFilterCondition(nameof(TeamPreferenceEntity.DigestFrequency), QueryComparisons.Equal, digestFrequency);
            var combinedFilter = TableQuery.CombineFilters(partitionKeyCondition, TableOperators.And, digestFrequencyCondition);

            TableQuery<TeamPreferenceEntity> query = new TableQuery<TeamPreferenceEntity>().Where(combinedFilter);
            TableContinuationToken continuationToken = null;
            var teamPreferenceCollection = new List<TeamPreferenceEntity>();

            do
            {
                var queryResult = await this.SubmitIdeaCloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);

                if (queryResult?.Results != null)
                {
                    teamPreferenceCollection.AddRange(queryResult.Results);
                    continuationToken = queryResult.ContinuationToken;
                }
            }
            while (continuationToken != null);

            return teamPreferenceCollection;
        }

        /// <summary>
        /// Stores or update team preference data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamPreferenceEntity">Represents team preference entity object.</param>
        /// <returns>A boolean that represents team preference entity is successfully saved/updated or not.</returns>
        public async Task<bool> UpsertTeamPreferenceAsync(TeamPreferenceEntity teamPreferenceEntity)
        {
            var result = await this.StoreOrUpdateEntityAsync(teamPreferenceEntity);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }

        /// <summary>
        /// Stores or update team preference data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="teamPreferenceEntity">Holds team preference detail entity data.</param>
        /// <returns>A task that represents team preference entity data is saved or updated.</returns>
        private async Task<TableResult> StoreOrUpdateEntityAsync(TeamPreferenceEntity teamPreferenceEntity)
        {
            await this.EnsureInitializedAsync();
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(teamPreferenceEntity);
            return await this.SubmitIdeaCloudTable.ExecuteAsync(addOrUpdateOperation);
        }
    }
}
