// <copyright file="CardHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Helpers
{
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.SubmitIdea;
    using Microsoft.Teams.Apps.SubmitIdea.Common;

    /// <summary>
    /// Class that handles the card helper methods.
    /// </summary>
    public static class CardHelper
    {
        /// <summary>
        ///  Represents the submit idea task module height.
        /// </summary>
        private const int SubmitIdeaTaskModuleHeight = 600;

        /// <summary>
        /// Represents the submit idea task module width.
        /// </summary>
        private const int SubmitIdeaTaskModuleWidth = 700;

        /// <summary>
        /// Sets the height of the task module.
        /// </summary>
        private const int ConfigurePreferencesTaskModuleHeight = 460;

        /// <summary>
        /// Sets the width of the task module.
        /// </summary>
        private const int ConfigurePreferencesTaskModuleWidth = 600;

        /// <summary>
        /// Get task module response.
        /// </summary>
        /// <param name="applicationBasePath">Represents the Application base Uri.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="command">Bot command text.</param>
        /// <returns>Returns task module response.</returns>
        public static TaskModuleResponse GetTaskModuleResponse(string applicationBasePath, IStringLocalizer<Strings> localizer, string command)
        {
            if (command.ToUpperInvariant() == Constants.SubmitAnIdeaAction)
            {
                return new TaskModuleResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Type = "continue",
                        Value = new TaskModuleTaskInfo()
                        {
                            Url = $"{applicationBasePath}/submit-idea",
                            Height = SubmitIdeaTaskModuleHeight,
                            Width = SubmitIdeaTaskModuleWidth,
                            Title = localizer.GetString("SubmitIdeaTitleText"),
                        },
                    },
                };
            }
            else
            {
                return new TaskModuleResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Type = "continue",
                        Value = new TaskModuleTaskInfo()
                        {
                            Url = $"{applicationBasePath}/configure-preferences",
                            Height = ConfigurePreferencesTaskModuleHeight,
                            Width = ConfigurePreferencesTaskModuleWidth,
                            Title = localizer.GetString("DigestPreferenceTitleText"),
                        },
                    },
                };
            }
        }
    }
}