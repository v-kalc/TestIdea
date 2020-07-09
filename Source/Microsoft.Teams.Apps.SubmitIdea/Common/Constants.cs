﻿// <copyright file="Constants.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Common
{
    /// <summary>
    /// A class that holds application constants that are used in multiple files.
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// Describes adaptive card version to be used. Version can be upgraded or changed using this value.
        /// </summary>
        public const string AdaptiveCardVersion = "1.2";

        /// <summary>
        /// Partition key for team post entity table.
        /// </summary>
        public const string TeamIdeaEntityPartitionKey = "TeamIdeaEntity";

        /// <summary>
        /// Partition key for team preference entity table.
        /// </summary>
        public const string TeamPreferenceEntityPartitionKey = "TeamPreferenceEntity";

        /// <summary>
        /// Partition key for team category entity table.
        /// </summary>
        public const string TeamCategoryEntityPartitionKey = "TeamCategoryEntity";

        /// <summary>
        /// Partition key for user team membership entity table.
        /// </summary>
        public const string UserTeamMembershipPartitionKey = "UserTeamMembershipEntity";

        /// <summary>
        /// All items post command id in the manifest file.
        /// </summary>
        public const string AllItemsPostCommandId = "allItems";

        /// <summary>
        /// All items post command id in the manifest file.
        /// </summary>
        public const string AllItemsIdeasCommandId = "ALLIDEAS";

        /// <summary>
        ///  Pending status command id in the manifest file.
        /// </summary>
        public const string PendingIdeaCommandId = "PENDING";

        /// <summary>
        ///  Approved status command id in the manifest file.
        /// </summary>
        public const string ApprovedIdeaCommandId = "APPROVED";

        /// <summary>
        /// Bot preference settings command to set preference for sending Weekly/Monthly notifications.
        /// </summary>
        public const string PreferenceSettings = "PREFERENCES";

        /// <summary>
        /// Partition key for team tag entity table.
        /// </summary>
        public const string TeamTagEntityPartitionKey = "TeamTagEntity";

        /// <summary>
        /// Bot Help command in personal scope.
        /// </summary>
        public const string HelpCommand = "HELP";

        /// <summary>
        /// Per page post count for lazy loading (max 50).
        /// </summary>
        public const int LazyLoadPerPagePostCount = 50;

        /// <summary>
        /// Submit an idea action.
        /// </summary>
        public const string SubmitAnIdeaAction = "SUBMIT AN IDEA";

        /// <summary>
        /// Preferences action.
        /// </summary>
        public const string Preferences = "PREFERENCES";

        /// <summary>
        /// default value for channel activity to send notifications.
        /// </summary>
        public const string TeamsBotFrameworkChannelId = "msteams";
    }
}
