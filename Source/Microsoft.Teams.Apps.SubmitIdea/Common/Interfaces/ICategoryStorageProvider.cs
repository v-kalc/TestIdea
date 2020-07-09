// <copyright file="ICategoryStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.SubmitIdea.Models;

    /// <summary>
    /// Category storage provider interface.
    /// </summary>
    public interface ICategoryStorageProvider
    {
        /// <summary>
        /// Add or update category in table storage.
        /// </summary>
        /// <param name="categoryEntity">represents the category entity that needs to be stored or updated.</param>
        /// <returns>category entity that is added or updated.</returns>
        Task<CategoryEntity> AddOrUpdateCategoryAsync(CategoryEntity categoryEntity);

        /// <summary>
        /// This method is used to delete categories for provided category Ids.
        /// </summary>
        /// <param name="categoryIds">list of categoryId that needs to be deleted.</param>
        /// <returns>boolean result.</returns>
        Task<bool> DeleteCategoriesAsync(IEnumerable<string> categoryIds);

        /// <summary>
        /// This method is used to get all categories.
        /// </summary>
        /// <returns>list of all category.</returns>
        Task<IEnumerable<CategoryEntity>> GetCategoriesAsync();

        /// <summary>
        /// This method is used to get category details by id.
        /// </summary>
        /// <param name="categoryIds">List of idea category id.</param>
        /// <returns>list of all category.</returns>
        Task<IEnumerable<CategoryEntity>> GetCategoriesByIdsAsync(IEnumerable<string> categoryIds);

        /// <summary>
        /// This method is used to fetch category details for a given category Id.
        /// </summary>
        /// <param name="categoryId">Category Id.</param>
        /// <returns>Category details.</returns>
        Task<CategoryEntity> GetCategoryDetailsAsync(string categoryId);
    }
}