// <copyright file="PolicyNames.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Authentication
{
    /// <summary>
    /// This class lists the names of the custom authorization policies in the project.
    /// </summary>
    public static class PolicyNames
    {
        /// <summary>
        /// The name of the authorization policy, MustBeTeamMemberUserPolicy.
        /// Indicates that user is a part of team and has permission to nominate and endorse team members.
        /// </summary>
        public const string MustBeTeamMemberUserPolicy = "MustBeTeamMemberUserPolicy";
    }
}
