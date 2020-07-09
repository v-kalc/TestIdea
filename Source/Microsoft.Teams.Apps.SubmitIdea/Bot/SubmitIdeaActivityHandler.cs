// <copyright file="SubmitIdeaActivityHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Bot
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.SubmitIdea.Cards;
    using Microsoft.Teams.Apps.SubmitIdea.Common;
    using Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces;
    using Microsoft.Teams.Apps.SubmitIdea.Helpers;
    using Microsoft.Teams.Apps.SubmitIdea.Models;
    using Microsoft.Teams.Apps.SubmitIdea.Models.Configuration;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// This class is responsible for reacting to incoming events from Microsoft Teams sent from BotFramework.
    /// </summary>
    public sealed class SubmitIdeaActivityHandler : TeamsActivityHandler
    {
        /// <summary>
        /// Represents the conversation type as personal.
        /// </summary>
        private const string Personal = "personal";

        /// <summary>
        /// Represents the conversation type as channel.
        /// </summary>
        private const string Channel = "channel";

        /// <summary>
        /// State management object for maintaining user conversation state.
        /// </summary>
        private readonly BotState userState;

        /// <summary>
        /// A set of key/value application configuration properties for Activity settings.
        /// </summary>
        private readonly IOptions<BotSettings> botOptions;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<SubmitIdeaActivityHandler> logger;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Instance of Application Insights Telemetry client.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Messaging Extension search helper for working with team posts data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IMessagingExtensionHelper messagingExtensionHelper;

        /// <summary>
        /// Instance of team preference storage helper.
        /// </summary>
        private readonly ITeamPreferenceStorageHelper teamPreferenceStorageHelper;

        /// <summary>
        /// Instance of team preference storage provider for team preferences.
        /// </summary>
        private readonly ITeamPreferenceStorageProvider teamPreferenceStorageProvider;

        /// <summary>
        /// Instance of team tags storage provider to configure team tags.
        /// </summary>
        private readonly ITeamTagStorageProvider teamTagStorageProvider;

        /// <summary>
        /// Provider for fetching information about team details from storage table.
        /// </summary>
        private readonly ITeamStorageProvider teamStorageProvider;

        /// <summary>
        /// A set of key/value application configuration properties for Activity settings.
        /// </summary>
        private readonly IOptions<SubmitIdeaActivityHandlerOptions> options;

        /// <summary>
        /// Represents unique id of a Team.
        /// </summary>
        private readonly string teamId;

        /// <summary>
        /// Initializes a new instance of the <see cref="SubmitIdeaActivityHandler"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        /// <param name="options">>A set of key/value application configuration properties for activity handler.</param>
        /// <param name="messagingExtensionHelper">Messaging Extension helper dependency injection.</param>
        /// <param name="userState">State management object for maintaining user conversation state.</param>
        /// <param name="teamPreferenceStorageHelper">Team preference storage helper dependency injection.</param>
        /// <param name="teamPreferenceStorageProvider">Team preference storage provider dependency injection.</param>
        /// <param name="teamTagStorageProvider">Team tags storage provider dependency injection.</param>
        /// <param name="botOptions">A set of key/value application configuration properties for activity handler.</param>
        /// <param name="teamStorageProvider">Provider for fetching information about team details from storage table.</param>
        public SubmitIdeaActivityHandler(
            ILogger<SubmitIdeaActivityHandler> logger,
            IStringLocalizer<Strings> localizer,
            TelemetryClient telemetryClient,
            IOptions<SubmitIdeaActivityHandlerOptions> options,
            IMessagingExtensionHelper messagingExtensionHelper,
            UserState userState,
            ITeamPreferenceStorageHelper teamPreferenceStorageHelper,
            ITeamPreferenceStorageProvider teamPreferenceStorageProvider,
            ITeamTagStorageProvider teamTagStorageProvider,
            IOptions<BotSettings> botOptions,
            ITeamStorageProvider teamStorageProvider)
        {
            this.logger = logger;
            this.localizer = localizer;
            this.telemetryClient = telemetryClient;
            options = options ?? throw new ArgumentNullException(nameof(options));
            this.messagingExtensionHelper = messagingExtensionHelper;
            this.userState = userState;
            this.teamPreferenceStorageHelper = teamPreferenceStorageHelper;
            this.teamPreferenceStorageProvider = teamPreferenceStorageProvider;
            this.teamTagStorageProvider = teamTagStorageProvider;
            this.botOptions = botOptions ?? throw new ArgumentNullException(nameof(botOptions));
            this.options = options;
            this.teamId = botOptions.Value.TeamId;
            this.teamStorageProvider = teamStorageProvider;
        }

        /// <summary>
        /// Handles an incoming activity.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.activityhandler.onturnasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        public override Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnTurnAsync), turnContext);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error at OnTurnAsync(): {ex.Message}", SeverityLevel.Error);
            }

            return base.OnTurnAsync(turnContext, cancellationToken);
        }

        /// <summary>
        /// Invoked when members other than this bot (like a user) are removed from the conversation.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnConversationUpdateActivityAsync(ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnConversationUpdateActivityAsync), turnContext);

                var activity = turnContext.Activity;
                this.logger.LogInformation($"conversationType: {activity.Conversation.ConversationType}, membersAdded: {activity.MembersAdded?.Count}, membersRemoved: {activity.MembersRemoved?.Count}");

                if (activity.Conversation.ConversationType.Equals(Personal, StringComparison.OrdinalIgnoreCase))
                {
                    if (activity.MembersAdded != null && activity.MembersAdded.Any(member => member.Id != activity.Recipient.Id))
                    {
                        await this.HandleMemberAddedinPersonalScopeAsync(turnContext);
                    }
                    else if (activity.MembersRemoved != null && activity.MembersRemoved.Any(member => member.Id != activity.Recipient.Id))
                    {
                        await this.HandleMemberRemovedInPersonalScopeAsync(turnContext);
                    }
                }
                else if (activity.Conversation.ConversationType.Equals(Channel, StringComparison.OrdinalIgnoreCase))
                {
                    if (activity.MembersAdded != null && activity.MembersAdded.Any(member => member.Id == activity.Recipient.Id))
                    {
                        await this.HandleMemberAddedInTeamAsync(turnContext);
                    }
                    else if (activity.MembersRemoved != null && activity.MembersRemoved.Any(member => member.Id == activity.Recipient.Id))
                    {
                        await this.HandleMemberRemovedInTeamScopeAsync(turnContext);
                    }
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Exception occurred while bot conversation update event.");
                throw;
            }
        }

        /// <summary>
        /// Invoked when the user opens the Messaging Extension or searching any content in it.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="query">Contains Messaging Extension query keywords.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Messaging extension response object to fill compose extension section.</returns>
        /// <remarks>
        /// https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionqueryasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(
            ITurnContext<IInvokeActivity> turnContext,
            MessagingExtensionQuery query,
            CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnTeamsMessagingExtensionQueryAsync), turnContext);

                var activity = turnContext.Activity;

                var messagingExtensionQuery = JsonConvert.DeserializeObject<MessagingExtensionQuery>(activity.Value.ToString());
                var searchQuery = this.messagingExtensionHelper.GetSearchResult(messagingExtensionQuery);

                return new MessagingExtensionResponse
                {
                    ComposeExtension = await this.messagingExtensionHelper.GetTeamPostSearchResultAsync(searchQuery, messagingExtensionQuery.CommandId, activity.From.AadObjectId, messagingExtensionQuery.QueryOptions.Count, messagingExtensionQuery.QueryOptions.Skip),
                };
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Failed to handle the Messaging Extension command {turnContext.Activity.Name}: {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Invoked when task module fetch event is received from the bot.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                taskModuleRequest = taskModuleRequest ?? throw new ArgumentNullException(nameof(taskModuleRequest));

                this.RecordEvent(nameof(this.OnTeamsTaskModuleFetchAsync), turnContext);

                var activity = turnContext.Activity;
                if (taskModuleRequest.Data == null)
                {
                    this.telemetryClient.TrackTrace("Request data obtained on task module fetch action is null.");
                    await turnContext.SendActivityAsync(this.localizer.GetString("WelcomeCardContent")).ConfigureAwait(false);
                    return default;
                }

                var postedValues = JsonConvert.DeserializeObject<BotCommand>(JObject.Parse(taskModuleRequest.Data.ToString()).SelectToken("data").ToString());
                var command = postedValues.Text;
                return CardHelper.GetTaskModuleResponse(this.options.Value.AppBaseUri, this.localizer, command);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while fetching task module received by the bot.");
                throw;
            }
        }

        /// <summary>
        /// Invoked when a message activity is received from the bot.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                var message = turnContext.Activity;

                message = message ?? throw new NullReferenceException(nameof(message));
                var command = message.RemoveRecipientMention().Trim();
                if (message.Conversation.ConversationType == Channel)
                {
                    switch (command.ToUpperInvariant())
                    {
                        case Constants.HelpCommand: // Help command to get the information about the bot.
                            this.logger.LogInformation("Sending user help card.");
                            var userHelpCards = CarouselCard.GetUserHelpCards(this.options.Value.AppBaseUri);
                            await turnContext.SendActivityAsync(MessageFactory.Carousel(userHelpCards)).ConfigureAwait(false);
                            break;
                        case Constants.PreferenceSettings: // Preference command to get the card to setup the tags preference of a team.
                            await turnContext.SendActivityAsync(MessageFactory.Attachment(WelcomeCard.GetPreferenceCard(localizer: this.localizer)), cancellationToken).ConfigureAwait(false);
                            break;
                        default:
                            await turnContext.SendActivityAsync(MessageFactory.Text(this.localizer.GetString("UnsupportedBotCommandText"))).ConfigureAwait(false);
                            this.logger.LogInformation($"Received a command {command.ToUpperInvariant()} which is not supported.");
                            break;
                    }
                }
                else
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text(this.localizer.GetString("UnsupportedBotPersonalCommandText"))).ConfigureAwait(false);
                    this.logger.LogInformation($"Received a command {command.ToUpperInvariant()} which is not supported.");
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while message activity is received from the bot.");
                throw;
            }
        }

        /// <summary>
        /// When OnTurn method receives a submit invoke activity on bot turn, it calls this method.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents a task module response.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleSubmitAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                taskModuleRequest = taskModuleRequest ?? throw new ArgumentNullException(nameof(taskModuleRequest));

                var preferenceData = JsonConvert.DeserializeObject<Preference>(taskModuleRequest.Data?.ToString());

                if (preferenceData == null)
                {
                    this.logger.LogInformation($"Request data obtained on task module submit action is null.");
                    await turnContext.SendActivityAsync(Strings.ErrorMessage).ConfigureAwait(false);
                    return null;
                }
                else if (preferenceData?.ConfigureDetails != null)
                {
                    var teamPreferenceDetail = this.teamPreferenceStorageHelper.CreateTeamPreferenceModel(preferenceData.ConfigureDetails);
                    await this.teamPreferenceStorageProvider.UpsertTeamPreferenceAsync(teamPreferenceDetail);
                }

                return null;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error in submit action of task module.");
                await turnContext.SendActivityAsync(Strings.ErrorMessage).ConfigureAwait(false);
                return null;
            }
        }

        /// <summary>
        /// Records event data to Application Insights telemetry client
        /// </summary>
        /// <param name="eventName">Name of the event.</param>
        /// <param name="turnContext">Provides context for a turn in a bot.</param>
        private void RecordEvent(string eventName, ITurnContext turnContext)
        {
            var teamsChannelData = turnContext.Activity.GetChannelData<TeamsChannelData>();

            this.telemetryClient.TrackEvent(eventName, new Dictionary<string, string>
            {
                { "userId", turnContext.Activity.From.AadObjectId },
                { "tenantId", turnContext.Activity.Conversation.TenantId },
                { "teamId", teamsChannelData?.Team?.Id },
                { "channelId", teamsChannelData?.Channel?.Id },
            });
        }

        private async Task<IEnumerable<TeamsChannelAccount>> GetTeamMembersAsync(ITurnContext<IConversationUpdateActivity> turnContext)
        {
            var teamInfo = turnContext.Activity.TeamsGetTeamInfo();
            return await TeamsInfo.GetTeamMembersAsync(turnContext, teamInfo.Id);
        }

        /// <summary>
        /// Get Azure Active Directory id of user.
        /// </summary>
        /// <param name="channelAccount">Channel account object.</param>
        /// <returns>Azure Active Directory id of user.</returns>
        private string GetUserAadObjectId(ChannelAccount channelAccount)
        {
            if (!string.IsNullOrWhiteSpace(channelAccount.AadObjectId))
            {
                return channelAccount.AadObjectId;
            }

            return channelAccount.Properties["objectId"].ToString();
        }

        /// <summary>
        /// Sent welcome card to personal chat.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn in a bot.</param>
        /// <returns>A task that represents a response.</returns>
        private async Task HandleMemberAddedinPersonalScopeAsync(ITurnContext<IConversationUpdateActivity> turnContext)
        {
            this.logger.LogInformation($"Bot added in personal {turnContext.Activity.Conversation.Id}");
            var userStateAccessors = this.userState.CreateProperty<UserConversationState>(nameof(UserConversationState));
            var userConversationState = await userStateAccessors.GetAsync(turnContext, () => new UserConversationState());

            userConversationState = userConversationState ?? throw new NullReferenceException(nameof(userConversationState));

            if (userConversationState.IsWelcomeCardSent)
            {
                return;
            }

            userConversationState.IsWelcomeCardSent = true;
            await userStateAccessors.SetAsync(turnContext, userConversationState);

            var userWelcomeCardAttachment = WelcomeCard.GetWelcomeCardAttachmentForPersonal(
                this.options.Value.AppBaseUri,
                localizer: this.localizer);

            await turnContext.SendActivityAsync(MessageFactory.Attachment(userWelcomeCardAttachment));
        }

        /// <summary>
        /// Set user conversation state to new if bot is removed from personal scope.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn in a bot.</param>
        /// <returns>>A task that represents a response.</returns>
        private async Task HandleMemberRemovedInPersonalScopeAsync(ITurnContext<IConversationUpdateActivity> turnContext)
        {
            this.logger.LogInformation($"Bot removed from personal {turnContext.Activity.Conversation.Id}");
            var userStateAccessors = this.userState.CreateProperty<UserConversationState>(nameof(UserConversationState));
            var userdata = await userStateAccessors.GetAsync(turnContext, () => new UserConversationState());
            userdata.IsWelcomeCardSent = false;
            await userStateAccessors.SetAsync(turnContext, userdata).ConfigureAwait(false);
        }

        /// <summary>
        /// Add user membership to storage if bot is installed in Team scope.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn in a bot.</param>
        /// <returns>A task that represents a response.</returns>
        private async Task HandleMemberAddedInTeamAsync(ITurnContext<IConversationUpdateActivity> turnContext)
        {
            this.logger.LogInformation($"Bot added in team {turnContext.Activity.Conversation.Id}");
            var channelData = turnContext.Activity.GetChannelData<TeamsChannelData>();
            var userWelcomeCardAttachment = channelData.Team.Id == this.teamId ? WelcomeCard.GetWelcomeCardAttachmentForCuratorTeam(this.options.Value.AppBaseUri, this.localizer) : WelcomeCard.GetWelcomeCardAttachmentForTeam(this.options.Value.AppBaseUri, this.localizer);

            // Storing team information to storage
            var teamsDetails = turnContext.Activity.TeamsGetTeamInfo();
            TeamEntity teamEntity = new TeamEntity
            {
                TeamId = teamsDetails.Id,
                BotInstalledOn = DateTime.UtcNow,
                ServiceUrl = turnContext.Activity.ServiceUrl,
                RowKey = teamsDetails.Id,
            };

            bool operationStatus = await this.teamStorageProvider.StoreOrUpdateTeamDetailAsync(teamEntity);
            if (!operationStatus)
            {
                this.logger.LogInformation($"Unable to store bot Installation detail in table storage.");
            }

            await turnContext.SendActivityAsync(MessageFactory.Attachment(userWelcomeCardAttachment));
        }

        /// <summary>
        /// Remove user membership from storage if bot is uninstalled from Team scope.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn in a bot.</param>
        /// <returns>A task that represents a response.</returns>
        private async Task HandleMemberRemovedInTeamScopeAsync(ITurnContext<IConversationUpdateActivity> turnContext)
        {
            this.logger.LogInformation($"Bot removed from team {turnContext.Activity.Conversation.Id}");
            var teamsChannelData = turnContext.Activity.GetChannelData<TeamsChannelData>();
            var teamId = teamsChannelData.Team.Id;

            // Deleting team information from storage when bot is uninstalled from a team.
            this.logger.LogInformation($"Bot removed {turnContext.Activity.Conversation.Id}");
            var teamEntity = await this.teamStorageProvider.GetTeamDetailAsync(teamId);
            bool operationStatus = await this.teamStorageProvider.DeleteTeamDetailAsync(teamEntity);
            if (!operationStatus)
            {
                this.logger.LogInformation($"Unable to remove team details from table storage.");
            }

            var result = await this.teamTagStorageProvider.DeleteTeamTagsEntryDataAsync(teamId);
            if (!result)
            {
                this.logger.LogInformation($"Filed to delete the tags for team: {teamId}");
            }
        }
    }
}