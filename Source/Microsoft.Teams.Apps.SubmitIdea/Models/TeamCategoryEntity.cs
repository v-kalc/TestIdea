// <copyright file="TeamCategoryEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using Microsoft.Teams.Apps.SubmitIdea.Common;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// A class that represents team category entity model.
    /// </summary>
    public class TeamCategoryEntity : TableEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TeamCategoryEntity"/> class.
        /// Holds team posts data.
        /// </summary>
        public TeamCategoryEntity()
        {
            this.PartitionKey = Constants.TeamCategoryEntityPartitionKey;
        }

        /// <summary>
        /// Gets or sets unique value for each Team where categories are configured.
        /// </summary>
        [Key]
        public string TeamId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets semicolon separated category Ids selected by user.
        /// </summary>
        [Required]
        public string Categories { get; set; }

        /// <summary>
        /// Gets or sets date time when entry is created by user in UTC format.
        /// </summary>
        public DateTime CreatedDate { get; set; }

        /// <summary>
        /// Gets or sets user name who configured tags in team.
        /// </summary>
        public string CreatedByName { get; set; }

        /// <summary>
        /// Gets or sets Azure Active Directory id of user who configured the categories in team.
        /// </summary>
        public string UserAadId { get; set; }
    }
}
