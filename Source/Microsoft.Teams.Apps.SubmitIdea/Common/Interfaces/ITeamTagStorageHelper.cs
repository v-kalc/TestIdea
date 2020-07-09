// <copyright file="ITeamTagStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces
{
    using Microsoft.Teams.Apps.SubmitIdea.Models;

    /// <summary>
    /// Interface for storage helper which helps in preparing model data for team tags configuration.
    /// </summary>
    public interface ITeamTagStorageHelper
    {
        /// <summary>
        /// Create team tags model data.
        /// </summary>
        /// <param name="teamTagEntity">Team tags detail.</param>
        /// <param name="userName">User name who has configured the tags in team.</param>
        /// <param name="userAadId">Azure Active Directory id of the user.</param>
        /// <returns>Represents team tag entity model.</returns>
        TeamTagEntity CreateTeamTagModel(TeamTagEntity teamTagEntity, string userName, string userAadId);
    }
}
