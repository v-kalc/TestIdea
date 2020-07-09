// <copyright file="DigestNotificationHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.SubmitIdea.Cards;
    using Microsoft.Teams.Apps.SubmitIdea.Common;
    using Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces;
    using Microsoft.Teams.Apps.SubmitIdea.Models;
    using Microsoft.Teams.Apps.SubmitIdea.Models.Configuration;
    using Polly;
    using Polly.Contrib.WaitAndRetry;
    using Polly.Retry;

    /// <summary>
    /// A class that handles sending notification to different channels.
    /// </summary>
    public class DigestNotificationHelper : IDigestNotificationHelper
    {
        /// <summary>
        /// Channel conversation type to send notification.
        /// </summary>
        private const string ChannelConversationType = "channel";

        /// <summary>
        /// Weekly digest for checking the digest notification type.
        /// </summary>
        private const string WeeklyDigest = "Weekly";

        /// <summary>
        /// Maximum no of ideas can be send to digest notification card.
        /// </summary>
        private const int MaxIdeasForNotification = 15;

        /// <summary>
        /// Retry policy with jitter, retry twice with a jitter delay of up to 1 sec. Retry for HTTP 429(transient error)/502 bad gateway.
        /// </summary>
        /// <remarks>
        /// Reference: https://github.com/Polly-Contrib/Polly.Contrib.WaitAndRetry#new-jitter-recommendation.
        /// </remarks>
        private readonly AsyncRetryPolicy retryPolicy = Policy.Handle<ErrorResponseException>(
            ex => ex.Response.StatusCode == HttpStatusCode.TooManyRequests || ex.Response.StatusCode == HttpStatusCode.InternalServerError)
            .WaitAndRetryAsync(Backoff.DecorrelatedJitterBackoffV2(TimeSpan.FromMilliseconds(1000), 2));

        /// <summary>
        /// Helper for storing channel details to azure table storage for sending notification.
        /// </summary>
        private readonly ITeamPreferenceStorageProvider teamPreferenceStorageProvider;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<DigestNotificationHelper> logger;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Bot adapter.
        /// </summary>
        private readonly IBotFrameworkHttpAdapter adapter;

        /// <summary>
        /// Represents a set of key/value application configuration properties for Share Now bot.
        /// </summary>
        private readonly IOptions<BotSettings> botOptions;

        /// <summary>
        /// A set of key/value application configuration properties for Activity settings.
        /// </summary>
        private readonly IOptions<SubmitIdeaActivityHandlerOptions> options;

        /// <summary>
        /// Instance of Search service for working with Microsoft Azure Table storage.
        /// </summary>
        private readonly ITeamIdeaSearchService teamIdeaSearchService;

        /// <summary>
        /// Instance of team idea storage helper to update post and get information of posts.
        /// </summary>
        private readonly ITeamIdeaStorageHelper teamIdeaStorageHelper;

        /// <summary>
        /// Instance of team storage provider to get information.
        /// </summary>
        private readonly ITeamStorageProvider teamStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="DigestNotificationHelper"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="botOptions">A set of key/value application configuration properties for Share Now bot.</param>
        /// <param name="adapter">Bot adapter.</param>
        /// <param name="teamPreferenceStorageProvider">Storage provider for team preference.</param>
        /// <param name="teamIdeaSearchService">The team idea search service dependency injection.</param>
        /// <param name="teamIdeaStorageHelper">Team idea storage helper dependency injection.</param>
        /// <param name="teamStorageProvider">Team storage provider dependency injection.</param>
        /// <param name="options">A set of key/value application configuration properties.</param>
        public DigestNotificationHelper(
            ILogger<DigestNotificationHelper> logger,
            IStringLocalizer<Strings> localizer,
            IOptions<BotSettings> botOptions,
            IBotFrameworkHttpAdapter adapter,
            ITeamPreferenceStorageProvider teamPreferenceStorageProvider,
            ITeamIdeaSearchService teamIdeaSearchService,
            ITeamIdeaStorageHelper teamIdeaStorageHelper,
            ITeamStorageProvider teamStorageProvider,
            IOptions<SubmitIdeaActivityHandlerOptions> options)
        {
            this.logger = logger;
            this.localizer = localizer;
            this.botOptions = botOptions ?? throw new ArgumentNullException(nameof(botOptions));
            this.adapter = adapter;
            this.teamPreferenceStorageProvider = teamPreferenceStorageProvider;
            this.teamIdeaSearchService = teamIdeaSearchService;
            this.teamIdeaStorageHelper = teamIdeaStorageHelper;
            this.teamStorageProvider = teamStorageProvider;
            this.options = options;
        }

        /// <summary>
        /// Send notification in channels on weekly or monthly basis as per the configured preference in different channels.
        /// Fetch data based on the date range and send it accordingly.
        /// </summary>
        /// <param name="startDate">Start date from which data should fetch.</param>
        /// <param name="endDate">End date till when data should fetch.</param>
        /// <param name="digestFrequency">Digest frequency text for notification like Monthly/Weekly.</param>
        /// <returns>A task that sends notification in channel.</returns>
        public async Task SendNotificationInChannelAsync(DateTime startDate, DateTime endDate, string digestFrequency)
        {
            try
            {
                digestFrequency = digestFrequency ?? throw new ArgumentNullException(nameof(digestFrequency));

                this.logger.LogInformation($"Send notification Timer trigger function executed at: {DateTime.UtcNow}");

                var teamPosts = await this.teamIdeaSearchService.GetTeamIdeasAsync(TeamPostSearchScope.FilterPostsAsPerDateRange, searchQuery: null, userObjectId: null);
                var filteredTeamPosts = this.teamIdeaStorageHelper.GetTeamIdeasInDateRangeAsync(teamPosts, startDate, endDate);

                if (filteredTeamPosts.Any())
                {
                    var teamPreferences = await this.teamPreferenceStorageProvider.GetTeamPreferencesAsync(digestFrequency);
                    var notificationCardTitle = digestFrequency.Equals(WeeklyDigest, StringComparison.InvariantCultureIgnoreCase)
                        ? this.localizer.GetString("NotificationCardWeeklyTitleText")
                        : this.localizer.GetString("NotificationCardMonthlyTitleText");

                    foreach (var teamPreference in teamPreferences)
                    {
                        var categoriesFilteredData = this.GetDataAsPerCategories(teamPreference, filteredTeamPosts);

                        if (categoriesFilteredData.Any())
                        {
                            var notificationCard = DigestNotificationListCard.GetNotificationListCard(
                                this.options.Value.AppBaseUri,
                                categoriesFilteredData,
                                notificationCardTitle);

                            var teamDetails = await this.teamStorageProvider.GetTeamDetailAsync(teamPreference.TeamId);
                            if (teamDetails != null)
                            {
                                await this.SendCardToTeamAsync(teamPreference, notificationCard, teamDetails.ServiceUrl);
                            }
                        }
                    }
                }
                else
                {
                    this.logger.LogInformation($"There is no digest data available to send at this time range from: {0} till {1}", startDate, endDate);
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while sending digest notifications.");
            }
        }

        /// <summary>
        /// Send the given attachment to the specified team.
        /// </summary>
        /// <param name="teamPreferenceEntity">Team preference model object.</param>
        /// <param name="cardToSend">The attachment card to send.</param>
        /// <param name="serviceUrl">Service url for a particular team.</param>
        /// <returns>A task that sends notification card in channel.</returns>
        private async Task SendCardToTeamAsync(
            TeamPreferenceEntity teamPreferenceEntity,
            Attachment cardToSend,
            string serviceUrl)
        {
            try
            {
                MicrosoftAppCredentials.TrustServiceUrl(serviceUrl);
                string teamsChannelId = teamPreferenceEntity.TeamId;

                var conversationReference = new ConversationReference()
                {
                    ChannelId = Constants.TeamsBotFrameworkChannelId,
                    Bot = new ChannelAccount() { Id = $"28:{this.botOptions.Value.MicrosoftAppId}" },
                    ServiceUrl = serviceUrl,
                    Conversation = new ConversationAccount() { ConversationType = ChannelConversationType, IsGroup = true, Id = teamsChannelId, TenantId = this.botOptions.Value.TenantId },
                };

                this.logger.LogInformation($"sending notification to channelId- {teamsChannelId}");

                // Retry it in addition to the original call.
                await this.retryPolicy.ExecuteAsync(async () =>
                {
                    try
                    {
                        await ((BotFrameworkAdapter)this.adapter).ContinueConversationAsync(
                        this.botOptions.Value.MicrosoftAppId,
                        conversationReference,
                        async (conversationTurnContext, conversationCancellationToken) =>
                        {
                            await conversationTurnContext.SendActivityAsync(MessageFactory.Attachment(cardToSend));
                        },
                        CancellationToken.None);
                    }
                    catch (Exception ex)
                    {
                        this.logger.LogError(ex, $"Error while performing retry logic to send digest notification to channel for team: {teamsChannelId}.");
                        throw;
                    }
                });
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while sending digest notification to channel from background service.");
            }
        }

        /// <summary>
        /// Get team posts as per configured categories for preference.
        /// </summary>
        /// <param name="teamPreferenceEntity">Team preference model object.</param>
        /// <param name="teamPosts">List of team posts.</param>
        /// <returns>List of team posts as per preference categories.</returns>
        private IEnumerable<TeamIdeaEntity> GetDataAsPerCategories(
            TeamPreferenceEntity teamPreferenceEntity,
            IEnumerable<TeamIdeaEntity> teamPosts)
        {
            try
            {
                var filteredPosts = new List<TeamIdeaEntity>();
                var preferenceCategoryIdsList = teamPreferenceEntity.Categories.Split(";").Where(category => !string.IsNullOrWhiteSpace(category)).ToList();
                teamPosts = teamPosts.OrderByDescending(c => c.UpdatedDate);

                // Loop through the list of filtered posts.
                foreach (var teamPost in teamPosts)
                {
                    // Loop through the list of preference category ids.
                    foreach (var preferenceCategoryId in preferenceCategoryIdsList)
                    {
                        if (teamPost.CategoryId == preferenceCategoryId && filteredPosts.Count < MaxIdeasForNotification)
                        {
                            // If preference category is present then add it in the list.
                            filteredPosts.Add(teamPost);
                            break; // break the inner loop to check for next post.
                        }
                    }

                    // Break the entire loop after getting top 15 posts.
                    if (filteredPosts.Count >= MaxIdeasForNotification)
                    {
                        break;
                    }
                }

                return filteredPosts;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while filtering the team posts as per the configured preference tags.");
                throw;
            }
        }
    }
}
