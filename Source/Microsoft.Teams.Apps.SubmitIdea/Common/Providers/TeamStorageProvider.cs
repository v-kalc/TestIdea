// <copyright file="TeamStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Common.Providers
{
    using System;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces;
    using Microsoft.Teams.Apps.SubmitIdea.Models;
    using Microsoft.Teams.Apps.SubmitIdea.Models.Configuration;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Awards storage provider.
    /// </summary>
    public class TeamStorageProvider : BaseStorageProvider, ITeamStorageProvider
    {
        private const string TeamConfigurationTable = "TeamConfiguration";

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamStorageProvider"/> class.
        /// </summary>
        /// <param name="options">A set of key/value application configuration properties for Microsoft Azure Table storage.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        public TeamStorageProvider(
            IOptions<StorageSettings> options,
            ILogger<CategoryStorageProvider> logger)
            : base(options?.Value.ConnectionString, TeamConfigurationTable, logger)
        {
        }

        /// <summary>
        /// Store or update team detail in Azure table storage.
        /// </summary>
        /// <param name="teamEntity">Represents team entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/> that represents team entity is saved or updated.</returns>
        public async Task<bool> StoreOrUpdateTeamDetailAsync(TeamEntity teamEntity)
        {
            await this.EnsureInitializedAsync();
            teamEntity = teamEntity ?? throw new ArgumentNullException(nameof(teamEntity));
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(teamEntity);
            var result = await this.SubmitIdeaCloudTable.ExecuteAsync(addOrUpdateOperation);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }

        /// <summary>
        /// Get already team detail from Azure table storage.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns><see cref="Task"/> Already saved team detail.</returns>
        public async Task<TeamEntity> GetTeamDetailAsync(string teamId)
        {
            await this.EnsureInitializedAsync();
            var teamEntity = new TeamEntity();

            var operation = TableOperation.Retrieve<TeamEntity>(teamId, teamId);
            var data = await this.SubmitIdeaCloudTable.ExecuteAsync(operation);
            teamEntity = data.Result as TeamEntity;

            return teamEntity;
        }

        /// <summary>
        /// This method delete the team detail record from table.
        /// </summary>
        /// <param name="teamEntity">Team configuration table entity.</param>
        /// <returns>A <see cref="Task"/> of type bool where true represents entity record is successfully deleted from table while false indicates failure in deleting data.</returns>
        public async Task<bool> DeleteTeamDetailAsync(TeamEntity teamEntity)
        {
            await this.EnsureInitializedAsync();
            TableOperation insertOrMergeOperation = TableOperation.Delete(teamEntity);
            TableResult result = await this.SubmitIdeaCloudTable.ExecuteAsync(insertOrMergeOperation);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }
    }
}
