// <copyright file="CategoryStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Common.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces;
    using Microsoft.Teams.Apps.SubmitIdea.Models;
    using Microsoft.Teams.Apps.SubmitIdea.Models.Configuration;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Category storage provider.
    /// </summary>
    public class CategoryStorageProvider : BaseStorageProvider, ICategoryStorageProvider
    {
        private const string CategoryTable = "Category";

        /// <summary>
        /// Initializes a new instance of the <see cref="CategoryStorageProvider"/> class.
        /// </summary>
        /// <param name="options">A set of key/value application configuration properties for Microsoft Azure Table storage.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        public CategoryStorageProvider(
            IOptions<StorageSettings> options,
            ILogger<CategoryStorageProvider> logger)
            : base(options?.Value.ConnectionString, CategoryTable, logger)
        {
        }

        /// <summary>
        /// This method is used to get all categories.
        /// </summary>
        /// <returns>list of all category.</returns>
        public async Task<IEnumerable<CategoryEntity>> GetCategoriesAsync()
        {
            await this.EnsureInitializedAsync();
            string filter = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, CategoryEntity.CategoryPartitionKey);
            var query = new TableQuery<CategoryEntity>().Where(filter);
            TableContinuationToken continuationToken = null;
            var categories = new List<CategoryEntity>();

            do
            {
                var queryResult = await this.SubmitIdeaCloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                categories.AddRange(queryResult?.Results);
                continuationToken = queryResult?.ContinuationToken;
            }
            while (continuationToken != null);

            return categories.OrderByDescending(category => category.Timestamp);
        }

        /// <summary>
        /// This method is used to fetch category details for a given category Id.
        /// </summary>
        /// <param name="categoryId">Category Id.</param>
        /// <returns>Category details.</returns>
        public async Task<CategoryEntity> GetCategoryDetailsAsync(string categoryId)
        {
            await this.EnsureInitializedAsync();
            var operation = TableOperation.Retrieve<CategoryEntity>(CategoryEntity.CategoryPartitionKey, categoryId);
            var category = await this.SubmitIdeaCloudTable.ExecuteAsync(operation);
            return category.Result as CategoryEntity;
        }

        /// <summary>
        /// This method is used to get category details by ids.
        /// </summary>
        /// <param name="categoryIds">List of idea category ids.</param>
        /// <returns>list of all category.</returns>
        public async Task<IEnumerable<CategoryEntity>> GetCategoriesByIdsAsync(IEnumerable<string> categoryIds)
        {
            categoryIds = categoryIds ?? throw new ArgumentNullException(nameof(categoryIds));
            await this.EnsureInitializedAsync();
            string categoriesCondition = this.CreateCategoriesFilter(categoryIds);

            TableQuery<CategoryEntity> query = new TableQuery<CategoryEntity>().Where(categoriesCondition);
            TableContinuationToken continuationToken = null;
            var categoryCollection = new List<CategoryEntity>();
            do
            {
                var queryResult = await this.SubmitIdeaCloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                if (queryResult?.Results != null)
                {
                    categoryCollection.AddRange(queryResult.Results);
                    continuationToken = queryResult.ContinuationToken;
                }
            }
            while (continuationToken != null);

            return categoryCollection;
        }

        /// <summary>
        /// Add or update category in table storage.
        /// </summary>
        /// <param name="categoryEntity">represents the category entity that needs to be stored or updated.</param>
        /// <returns>category entity that is added or updated.</returns>
        public async Task<CategoryEntity> AddOrUpdateCategoryAsync(CategoryEntity categoryEntity)
        {
            await this.EnsureInitializedAsync();
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(categoryEntity);
            var result = await this.SubmitIdeaCloudTable.ExecuteAsync(addOrUpdateOperation);
            return result.Result as CategoryEntity;
        }

        /// <summary>
        /// This method is used to delete categories for provided category Ids.
        /// </summary>
        /// <param name="categoryIds">list of categoryId that needs to be deleted.</param>
        /// <returns>boolean result.</returns>
        public async Task<bool> DeleteCategoriesAsync(IEnumerable<string> categoryIds)
        {
            if (categoryIds == null)
            {
                throw new ArgumentNullException(nameof(categoryIds));
            }

            await this.EnsureInitializedAsync();

            foreach (var categoryId in categoryIds)
            {
                var operation = TableOperation.Retrieve<CategoryEntity>(CategoryEntity.CategoryPartitionKey, categoryId);
                var data = await this.SubmitIdeaCloudTable.ExecuteAsync(operation);
                var category = data.Result as CategoryEntity;
                if (category != null)
                {
                    TableOperation deleteOperation = TableOperation.Delete(category);
                    var result = await this.SubmitIdeaCloudTable.ExecuteAsync(deleteOperation);
                }
            }

            return true;
        }

        /// <summary>
        /// Get combined filter condition for user private posts data.
        /// </summary>
        /// <param name="categoryIds">List of user private posts id.</param>
        /// <returns>Returns combined filter for user private posts.</returns>
        private string CreateCategoriesFilter(IEnumerable<string> categoryIds)
        {
            var categoryIdConditions = new List<string>();
            StringBuilder combinedCaregoryIdsFilter = new StringBuilder();

            categoryIds = categoryIds.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct();

            foreach (var postId in categoryIds)
            {
                categoryIdConditions.Add("(" + TableQuery.GenerateFilterCondition("RowKey", QueryComparisons.Equal, postId) + ")");
            }

            if (categoryIdConditions.Count >= 2)
            {
                var categories = categoryIdConditions.Take(categoryIdConditions.Count - 1).ToList();

                categories.ForEach(postCondition =>
                {
                    combinedCaregoryIdsFilter.Append($"{postCondition} {"or"} ");
                });

                combinedCaregoryIdsFilter.Append($"{categoryIdConditions.Last()}");

                return combinedCaregoryIdsFilter.ToString();
            }
            else
            {
                return TableQuery.GenerateFilterCondition("RowKey", QueryComparisons.Equal, categoryIds.FirstOrDefault());
            }
        }
    }
}
