// <copyright file="SubmitIdeaActivityHandlerOptions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
namespace Microsoft.Teams.Apps.SubmitIdea.Models.Configuration
{
    /// <summary>
    /// This class provide options for the <see cref="SubmitIdeaActivityHandlerOptions" /> bot.
    /// </summary>
    public sealed class SubmitIdeaActivityHandlerOptions
    {
        /// <summary>
        /// Gets or sets application base URL used to return success or failure task module result.
        /// </summary>
        public string AppBaseUri { get; set; }
    }
}
