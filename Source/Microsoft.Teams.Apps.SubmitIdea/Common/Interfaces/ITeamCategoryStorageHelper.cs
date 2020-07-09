// <copyright file="ITeamCategoryStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces
{
    using Microsoft.Teams.Apps.SubmitIdea.Models;

    /// <summary>
    /// Interface for storage helper which helps in preparing model data for team categories configuration.
    /// </summary>
    public interface ITeamCategoryStorageHelper
    {
        /// <summary>
        /// Create team categories model data.
        /// </summary>
        /// <param name="teamCategoryEntity">Team categories detail.</param>
        /// <param name="userName">User name who has configured the categories in team.</param>
        /// <param name="userAadId">Azure Active Directory id of the user.</param>
        /// <returns>Represents team tag entity model.</returns>
        TeamCategoryEntity CreateTeamCategoryModel(TeamCategoryEntity teamCategoryEntity, string userName, string userAadId);
    }
}
