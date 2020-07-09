// <copyright file="CategoryController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces;
    using Microsoft.Teams.Apps.SubmitIdea.Models;

    /// <summary>
    /// This endpoint is used to manage categories.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class CategoryController : BaseSubmitIdeaController
    {
        /// <summary>
        /// Event name for team category HTTP post call.
        /// </summary>
        private const string RecordCategoryHTTPPostCall = "Categories - HTTP Post call succeeded";

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private ILogger<CategoryController> logger;

        /// <summary>
        /// Provider for managing categories from azure table storage.
        /// </summary>
        private ICategoryStorageProvider storageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="CategoryController"/> class.
        /// </summary>
        /// <param name="storageProvider">storageProvider.</param>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        public CategoryController(
            ICategoryStorageProvider storageProvider,
            ILogger<CategoryController> logger,
            TelemetryClient telemetryClient)
            : base(telemetryClient)
        {
            this.logger = logger;
            this.storageProvider = storageProvider;
        }

        /// <summary>
        /// This method is used to get all categories.
        /// </summary>
        /// <returns>categories.</returns>
        [HttpGet("allcategories")]
        public async Task<IActionResult> GetCategoriesAsync()
        {
            try
            {
                var categories = await this.storageProvider.GetCategoriesAsync();
                return this.Ok(categories);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "failed to get categories");
                throw;
            }
        }

        /// <summary>
        /// Post call to add or edit category in table storage.
        /// </summary>
        /// <param name="categoryEntity">category entity to be added or updated.</param>
        /// <returns>category entity.</returns>
        [HttpPost("category")]
        public async Task<IActionResult> PostAsync([FromBody] CategoryEntity categoryEntity)
        {
            try
            {
                if (string.IsNullOrEmpty(categoryEntity?.CategoryName))
                {
                    this.logger.LogError("Empty category name while creating category in Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Empty category name while creating category in Microsoft Azure Table storage.");
                }

                this.RecordEvent(RecordCategoryHTTPPostCall);
                if (categoryEntity.CategoryId == null)
                {
                    var newCategoryEntity = new CategoryEntity
                    {
                        CategoryId = Guid.NewGuid().ToString(),
                        CreatedOn = DateTime.UtcNow,
                        CreatedByUserId = this.UserAadId,
                        CategoryName = categoryEntity.CategoryName,
                        CategoryDescription = categoryEntity.CategoryDescription,
                    };

                    return this.Ok(await this.storageProvider.AddOrUpdateCategoryAsync(newCategoryEntity));
                }
                else
                {
                    var category = this.storageProvider.GetCategoryDetailsAsync(categoryEntity.CategoryId);
                    if (category == null)
                    {
                        this.logger.LogError($"User {this.UserAadId} is forbidden to update category {categoryEntity.CategoryId}.");
                        this.RecordEvent("Update idea - HTTP Put call failed");
                        return this.Forbid($"You do not have required access to update category {categoryEntity.CategoryId}.");
                    }

                    categoryEntity.ModifiedByUserId = this.UserAadId;
                    return this.Ok(await this.storageProvider.AddOrUpdateCategoryAsync(categoryEntity));
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while adding/updating the category.");
                throw;
            }
        }

        /// <summary>
        /// Delete call to delete categories for provided category Id.
        /// </summary>
        /// <param name="categoryIds">category Ids that needs to be deleted.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpDelete("categories")]
        public async Task<IActionResult> DeleteAsync(string categoryIds)
        {
            try
            {
                if (string.IsNullOrEmpty(categoryIds))
                {
                    this.logger.LogError("Error while deleting categories in Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Empty category Ids");
                }

                IList<string> categories = categoryIds?.Split(",");
                return this.Ok(await this.storageProvider.DeleteCategoriesAsync(categories));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while deleting categories.");
                throw;
            }
        }
    }
}
