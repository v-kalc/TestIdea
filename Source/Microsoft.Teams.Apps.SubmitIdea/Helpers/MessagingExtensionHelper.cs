// <copyright file="MessagingExtensionHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Web;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.SubmitIdea;
    using Microsoft.Teams.Apps.SubmitIdea.Common;
    using Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces;
    using Microsoft.Teams.Apps.SubmitIdea.Models;
    using Microsoft.Teams.Apps.SubmitIdea.Models.Configuration;

    /// <summary>
    /// A class that handles the search activities for Messaging Extension.
    /// </summary>
    public class MessagingExtensionHelper : IMessagingExtensionHelper
    {
        /// <summary>
        /// Search text parameter name in the manifest file.
        /// </summary>
        private const string SearchTextParameterName = "searchText";

        /// <summary>
        /// Maximum length for thumbnail card text.
        /// </summary>
        private const int ThumbnailCardTextMaxLength = 25;

        /// <summary>
        /// Maximum length for adaptive text block card text.
        /// </summary>
        private const int AdaptiveTextBlockTextMaxLength = 19;

        /// <summary>
        /// Instance of Search service for working with Microsoft Azure Table storage.
        /// </summary>
        private readonly ITeamIdeaSearchService teamIdeaSearchService;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// A set of key/value application configuration properties for Activity settings.
        /// </summary>
        private readonly IOptions<SubmitIdeaActivityHandlerOptions> options;

        /// <summary>
        /// Instance of idea category storage provider.
        /// </summary>
        private readonly ICategoryStorageProvider categoryStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="MessagingExtensionHelper"/> class.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="teamIdeaSearchService">The team post search service dependency injection.</param>
        /// <param name="categoryStorageProvider">The idea category storage provider.</param>
        /// <param name="options">>A set of key/value application configuration properties for activity handler.</param>
        public MessagingExtensionHelper(
            IStringLocalizer<Strings> localizer,
            ITeamIdeaSearchService teamIdeaSearchService,
            ICategoryStorageProvider categoryStorageProvider,
            IOptions<SubmitIdeaActivityHandlerOptions> options)
        {
            this.localizer = localizer;
            this.teamIdeaSearchService = teamIdeaSearchService;
            this.categoryStorageProvider = categoryStorageProvider;
            this.options = options;
        }

        /// <summary>
        /// Get the results from Azure Search service and populate the result (card + preview).
        /// </summary>
        /// <param name="query">Query which the user had typed in Messaging Extension search field.</param>
        /// <param name="commandId">Command id to determine which tab in Messaging Extension has been invoked.</param>
        /// <param name="userObjectId">Azure Active Directory id of the user.</param>
        /// <param name="count">Number of search results to return.</param>
        /// <param name="skip">Number of search results to skip.</param>
        /// <returns><see cref="Task"/>Returns Messaging Extension result object, which will be used for providing the card.</returns>
        public async Task<MessagingExtensionResult> GetTeamPostSearchResultAsync(
            string query,
            string commandId,
            string userObjectId,
            int? count,
            int? skip)
        {
            MessagingExtensionResult composeExtensionResult = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = AttachmentLayoutTypes.List,
                Attachments = new List<MessagingExtensionAttachment>(),
            };

            IEnumerable<TeamIdeaEntity> teamPostResults;

            // commandId should be equal to Id mentioned in Manifest file under composeExtensions section.
            switch (commandId?.ToUpperInvariant())
            {
                case Constants.AllItemsIdeasCommandId: // Get all ideas
                    teamPostResults = await this.teamIdeaSearchService.GetTeamIdeasAsync(TeamPostSearchScope.AllItems, query, userObjectId, count, skip);
                    composeExtensionResult = await this.GetTeamPostResultAsync(teamPostResults);
                    break;

                case Constants.PendingIdeaCommandId: // Get pending ideas.
                    teamPostResults = await this.teamIdeaSearchService.GetTeamIdeasAsync(TeamPostSearchScope.Pending, query, userObjectId, count, skip);
                    composeExtensionResult = await this.GetTeamPostResultAsync(teamPostResults);
                    break;

                case Constants.ApprovedIdeaCommandId: // Get approved ideas.
                    teamPostResults = await this.teamIdeaSearchService.GetTeamIdeasAsync(TeamPostSearchScope.Approved, query, userObjectId, count, skip);
                    composeExtensionResult = await this.GetTeamPostResultAsync(teamPostResults);
                    break;
            }

            return composeExtensionResult;
        }

        /// <summary>
        /// Get the value of the searchText parameter in the Messaging Extension query.
        /// </summary>
        /// <param name="query">Contains Messaging Extension query keywords.</param>
        /// <returns>A value of the searchText parameter.</returns>
        public string GetSearchResult(MessagingExtensionQuery query)
        {
            return query?.Parameters.FirstOrDefault(parameter => parameter.Name.Equals(SearchTextParameterName, StringComparison.OrdinalIgnoreCase))?.Value?.ToString();
        }

        /// <summary>
        /// Get team posts result for Messaging Extension.
        /// </summary>
        /// <param name="teamIdeaResults">List of user search result.</param>
        /// <returns><see cref="Task"/>Returns Messaging Extension result object, which will be used for providing the card.</returns>
        private async Task<MessagingExtensionResult> GetTeamPostResultAsync(IEnumerable<TeamIdeaEntity> teamIdeaResults)
        {
            MessagingExtensionResult composeExtensionResult = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = AttachmentLayoutTypes.List,
                Attachments = new List<MessagingExtensionAttachment>(),
            };

            if (teamIdeaResults == null)
            {
                return composeExtensionResult;
            }

            var catagoryDetails = await this.categoryStorageProvider.GetCategoriesByIdsAsync(teamIdeaResults.Select(teamIdea => teamIdea.CategoryId));

            foreach (var teamIdea in teamIdeaResults)
            {
                var card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 2))
                {
                    Body = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Text = teamIdea.Title,
                            Wrap = true,
                            Weight = AdaptiveTextWeight.Bolder,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = teamIdea.Description,
                            Wrap = true,
                            Size = AdaptiveTextSize.Small,
                        },
                    },
                };

                var categoryName = catagoryDetails.Where(categoryDetail => categoryDetail.CategoryId == teamIdea.CategoryId).FirstOrDefault()?.CategoryName;

                card.Body.Add(this.GetTagsContainer(teamIdea));
                card.Body.Add(this.GetPostTypeContainer(teamIdea, categoryName));

                var voteIcon = $"<img src='{this.options.Value.AppBaseUri}/Artifacts/voteIcon.png' alt='vote logo' width='18' height='18'";
                var nameString = teamIdea.CreatedByName.Length < ThumbnailCardTextMaxLength
                    ? HttpUtility.HtmlEncode(teamIdea.CreatedByName)
                    : $"{HttpUtility.HtmlEncode(teamIdea.CreatedByName.Substring(0, ThumbnailCardTextMaxLength - 1))} {"..."}";

                ThumbnailCard previewCard = new ThumbnailCard
                {
                    Title = $"<p style='font-weight: 600;'>{teamIdea.Title}</p>",
                    Text = $"{nameString} {"|"} {categoryName} {"|"} {teamIdea.TotalVotes} {voteIcon}",
                };

                composeExtensionResult.Attachments.Add(new Attachment
                {
                    ContentType = AdaptiveCard.ContentType,
                    Content = card,
                }.ToMessagingExtensionAttachment(previewCard.ToAttachment()));
            }

            return composeExtensionResult;
        }

        /// <summary>
        /// Get container for team ideas.
        /// </summary>
        /// <param name="teamIdea">Team post entity object.</param>
        /// <param name="categoryName">Name of current category.</param>
        /// <returns>Return a container for team ideas.</returns>
        private AdaptiveContainer GetPostTypeContainer(TeamIdeaEntity teamIdea, string categoryName)
        {
            string applicationBasePath = this.options.Value.AppBaseUri;

            var postTypeContainer = new AdaptiveContainer
            {
                Items = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveImage
                                    {
                                        Url = new Uri($"{applicationBasePath}/Artifacts/peopleAvatar.png"),
                                        Size = AdaptiveImageSize.Auto,
                                        Style = AdaptiveImageStyle.Person,
                                        AltText = "User Image",
                                    },
                                },
                                Spacing = AdaptiveSpacing.Small,
                            },
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = teamIdea.CreatedByName.Length > AdaptiveTextBlockTextMaxLength ? $"{teamIdea.CreatedByName.Substring(0, AdaptiveTextBlockTextMaxLength - 1)} {"..."}" : teamIdea.CreatedByName,
                                        Wrap = true,
                                    },
                                },
                                Spacing = AdaptiveSpacing.Small,
                            },
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveImage
                                    {
                                        Url = new Uri($"{this.options.Value.AppBaseUri}/Artifacts/videoTypeDot.png"),
                                        Size = AdaptiveImageSize.Stretch,
                                        Style = AdaptiveImageStyle.Default,
                                        Height = AdaptiveHeight.Auto,
                                    },
                                },
                                Spacing = AdaptiveSpacing.Small,
                            },
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = categoryName,
                                        Spacing = AdaptiveSpacing.None,
                                        IsSubtle = true,
                                        Wrap = true,
                                        Weight = AdaptiveTextWeight.Bolder,
                                    },
                                },
                                Spacing = AdaptiveSpacing.Small,
                            },
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = $"{teamIdea.TotalVotes} ",
                                        Wrap = true,
                                    },
                                },
                                Spacing = AdaptiveSpacing.Small,
                            },
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveImage
                                    {
                                        Url = new Uri($"{applicationBasePath}/Artifacts/voteIcon.png"),
                                        Size = AdaptiveImageSize.Stretch,
                                        Style = AdaptiveImageStyle.Default,
                                    },
                                },
                                Spacing = AdaptiveSpacing.Small,
                            },
                        },
                    },
                },
            };

            return postTypeContainer;
        }

        /// <summary>
        /// Get tags container for team post.
        /// </summary>
        /// <param name="teamIdea">Team post entity object.</param>
        /// <returns>Return a container for team post tags.</returns>
        private AdaptiveContainer GetTagsContainer(TeamIdeaEntity teamIdea)
        {
            var tagsContainer = new AdaptiveContainer
            {
                Items = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = $"**{this.localizer.GetString("TagsLabelText")}{":"}** {teamIdea.Tags?.Replace(";", ", ", false, CultureInfo.InvariantCulture)}",
                                        Wrap = true,
                                    },
                                },
                            },
                        },
                    },
                },
            };

            return tagsContainer;
        }
    }
}
