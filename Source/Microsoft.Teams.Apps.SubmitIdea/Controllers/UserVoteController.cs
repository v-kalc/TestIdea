// <copyright file="UserVoteController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Controllers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.SubmitIdea.Common.Interfaces;
    using Microsoft.Teams.Apps.SubmitIdea.Models;
    using Microsoft.WindowsAzure.Storage;
    using Polly;
    using Polly.Contrib.WaitAndRetry;
    using Polly.Retry;

    /// <summary>
    /// Controller to handle user vote operations.
    /// </summary>
    [ApiController]
    [Route("api/uservote")]
    [Authorize]
    public class UserVoteController : BaseSubmitIdeaController
    {
        /// <summary>
        /// Event name for user vote HTTP get call.
        /// </summary>
        private const string RecordUserVoteHTTPGetCall = "User votes - HTTP Get call succeeded.";

        /// <summary>
        /// Retry policy with jitter.
        /// </summary>
        /// <remarks>
        /// Reference: https://github.com/Polly-Contrib/Polly.Contrib.WaitAndRetry#new-jitter-recommendation.
        /// </remarks>
        private readonly AsyncRetryPolicy retryPolicy;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Instance of user vote storage provider to add and delete user vote.
        /// </summary>
        private readonly IUserVoteStorageProvider userVoteStorageProvider;

        /// <summary>
        /// Instance of team post storage provider.
        /// </summary>
        private readonly ITeamIdeaStorageProvider teamIdeaStorageProvider;

        /// <summary>
        /// Instance of Search service for working with Microsoft Azure Table storage.
        /// </summary>
        private readonly ITeamIdeaSearchService teamIdeaSearchService;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserVoteController"/> class.
        /// </summary>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        /// <param name="userVoteStorageProvider">Instance of user vote storage provider to add and delete user vote.</param>
        /// <param name="teamIdeaStorageProvider">Instance of team post storage provider to update post and get information of posts.</param>
        /// <param name="teamIdeaSearchService">The team post search service dependency injection.</param>
        public UserVoteController(
            ILogger<UserVoteController> logger,
            TelemetryClient telemetryClient,
            IUserVoteStorageProvider userVoteStorageProvider,
            ITeamIdeaStorageProvider teamIdeaStorageProvider,
            ITeamIdeaSearchService teamIdeaSearchService)
            : base(telemetryClient)
        {
            this.logger = logger;
            this.userVoteStorageProvider = userVoteStorageProvider;
            this.teamIdeaStorageProvider = teamIdeaStorageProvider;
            this.teamIdeaSearchService = teamIdeaSearchService;
            this.retryPolicy = Policy.Handle<StorageException>(ex => ex.RequestInformation.HttpStatusCode == StatusCodes.Status412PreconditionFailed)
               .WaitAndRetryAsync(Backoff.LinearBackoff(TimeSpan.FromMilliseconds(1000), 3));
        }

        /// <summary>
        /// Get call to retrieve list of votes for user.
        /// </summary>
        /// <returns>List of team posts.</returns>
        [HttpGet("votes")]
        public async Task<IActionResult> GetVotesAsync()
        {
            try
            {
                this.logger.LogInformation("call to retrieve list of votes for user.");

                var userVotes = await this.userVoteStorageProvider.GetVotesAsync(this.UserAadId);
                this.RecordEvent(RecordUserVoteHTTPGetCall);

                return this.Ok(userVotes);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to team post service.");
                throw;
            }
        }

        /// <summary>
        /// Stores user vote for a post.
        /// </summary>
        /// <param name="postCreatedByUserId">AAD user Id of user who created post.</param>
        /// <param name="postId">Id of the post to delete vote.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpGet("vote")]
        public async Task<IActionResult> AddVoteAsync(string postCreatedByUserId, string postId)
        {
            this.logger.LogInformation("call to add user vote.");

#pragma warning disable CA1062 // post details are validated by model validations for null check and is responded with bad request status
            var userVoteForPost = await this.userVoteStorageProvider.GetUserVoteForPostAsync(this.UserAadId, postId);
#pragma warning restore CA1062 // post details are validated by model validations for null check and is responded with bad request status

            if (userVoteForPost == null)
            {
                UserVoteEntity userVote = new UserVoteEntity
                {
                    UserId = this.UserAadId,
                    PostId = postId,
                };

                TeamIdeaEntity postEntity = null;
                bool isPostSavedSuccessful = false;

                // Retry if storage operation conflict occurs during updating user vote count.
                await this.retryPolicy.ExecuteAsync(async () =>
                {
                    try
                    {
                        postEntity = await this.teamIdeaStorageProvider.GetPostAsync(postCreatedByUserId, userVote.PostId);

                        // increment the vote count
                        // if the execution is retried, then get the latest vote count and increase it by 1
                        postEntity.TotalVotes += 1;

                        isPostSavedSuccessful = await this.teamIdeaStorageProvider.UpsertTeamIdeaAsync(postEntity);
                    }
                    catch (StorageException ex)
                    {
                        if (ex.RequestInformation.HttpStatusCode == StatusCodes.Status412PreconditionFailed)
                        {
                            this.logger.LogError("Optimistic concurrency violation – entity has changed since it was retrieved.");
                            throw;
                        }
                    }
#pragma warning disable CA1031 // catching generic exception to trace log error in telemetry and continue the execution
                    catch (Exception ex)
#pragma warning restore CA1031 // catching generic exception to trace log error in telemetry and continue the execution
                    {
                        // log exception details to telemetry
                        // but do not attempt to retry in order to avoid multiple vote count increment
                        this.logger.LogError(ex, "Exception occurred while reading post details.");
                    }
                });

                if (!isPostSavedSuccessful)
                {
                    this.logger.LogError($"Vote is not updated successfully for post {postId} by {this.UserAadId} ");
                    return this.StatusCode(StatusCodes.Status500InternalServerError, "Vote is not updated successfully.");
                }

                bool isUserVoteSavedSuccessful = false;

                this.logger.LogInformation($"Post vote count updated for PostId:{postId}");
                isUserVoteSavedSuccessful = await this.userVoteStorageProvider.UpsertUserVoteAsync(userVote);

                // if user vote is not saved successfully
                // revert back the total post count
                if (!isUserVoteSavedSuccessful)
                {
                    await this.retryPolicy.ExecuteAsync(async () =>
                    {
                        try
                        {
                            postEntity = await this.teamIdeaStorageProvider.GetPostAsync(postCreatedByUserId, userVote.PostId);
                            postEntity.TotalVotes -= 1;

                            // Update operation will throw exception if the column has already been updated
                            // or if there is a transient error (handled by an Azure storage)
                            await this.teamIdeaStorageProvider.UpsertTeamIdeaAsync(postEntity);
                            await this.teamIdeaSearchService.RunIndexerOnDemandAsync();
                        }
                        catch (StorageException ex)
                        {
                            if (ex.RequestInformation.HttpStatusCode == StatusCodes.Status412PreconditionFailed)
                            {
                                this.logger.LogError("Optimistic concurrency violation – entity has changed since it was retrieved.");
                                throw;
                            }
                        }
#pragma warning disable CA1031 // catching generic exception to trace log error in telemetry and continue the execution
                        catch (Exception ex)
#pragma warning restore CA1031 // catching generic exception to trace log error in telemetry and continue the execution
                        {
                            // log exception details to telemetry
                            // but do not attempt to retry in order to avoid multiple vote count decrement
                            this.logger.LogError(ex, "Exception occurred while reading post details.");
                        }
                    });
                }
                else
                {
                    this.logger.LogInformation($"User vote added for user{this.UserAadId} for PostId:{postId}");
                    await this.teamIdeaSearchService.RunIndexerOnDemandAsync();
                    return this.Ok(true);
                }
            }

            return this.Ok(false);
        }

        /// <summary>
        /// Deletes user vote for a post.
        /// </summary>
        /// <param name="postCreatedByUserId">AAD user Id of user who created post.</param>
        /// <param name="postId">Id of the post to delete vote.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpDelete]
        public async Task<IActionResult> DeleteVoteAsync(string postCreatedByUserId, string postId)
        {
            this.logger.LogInformation("call to delete user vote.");

            if (string.IsNullOrEmpty(postCreatedByUserId))
            {
                this.logger.LogError("Error while deleting vote. Parameter postCreatedByuserId is either null or empty.");
                return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while adding vote. Parameter postCreatedByuserId is either null or empty.");
            }

            if (string.IsNullOrEmpty(postId))
            {
                this.logger.LogError("Error while deleting vote. PostId is either null or empty.");
                return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while adding vote. PostId is either null or empty.");
            }

#pragma warning disable CA1062 // post details are validated by model validations for null check and is responded with bad request status
            var userVoteForPost = await this.userVoteStorageProvider.GetUserVoteForPostAsync(this.UserAadId, postId);
#pragma warning restore CA1062 // post details are validated by model validations for null check and is responded with bad request status

            if (userVoteForPost != null)
            {
                TeamIdeaEntity postEntity = null;
                bool isPostSavedSuccessful = false;

                // Retry if storage operation conflict occurs during updating user vote count.
                await this.retryPolicy.ExecuteAsync(async () =>
                {
                    try
                    {
                        postEntity = await this.teamIdeaStorageProvider.GetPostAsync(postCreatedByUserId, postId);

                        // increment the vote count
                        // if the execution is retried, then get the latest vote count and increase it by 1
                        postEntity.TotalVotes -= 1;

                        isPostSavedSuccessful = await this.teamIdeaStorageProvider.UpsertTeamIdeaAsync(postEntity);
                    }
                    catch (StorageException ex)
                    {
                        if (ex.RequestInformation.HttpStatusCode == StatusCodes.Status412PreconditionFailed)
                        {
                            this.logger.LogError("Optimistic concurrency violation – entity has changed since it was retrieved.");
                            throw;
                        }
                    }
#pragma warning disable CA1031 // catching generic exception to trace log error in telemetry and continue the execution
                    catch (Exception ex)
#pragma warning restore CA1031 // catching generic exception to trace log error in telemetry and continue the execution
                    {
                        // log exception details to telemetry
                        // but do not attempt to retry in order to avoid multiple vote count increment
                        this.logger.LogError(ex, "Exception occurred while reading post details.");
                    }
                });

                if (!isPostSavedSuccessful)
                {
                    this.logger.LogError($"Vote is not updated successfully for post {postId} by {postCreatedByUserId} ");
                    return this.StatusCode(StatusCodes.Status500InternalServerError, "Vote is not updated successfully.");
                }

                bool isUserVotDeletedSuccessful = false;

                this.logger.LogInformation($"Post vote count updated for PostId:{postId}");
                isUserVotDeletedSuccessful = await this.userVoteStorageProvider.DeleteEntityAsync(postId, postCreatedByUserId);

                // if user vote is not saved successfully
                // revert back the total post count
                if (!isUserVotDeletedSuccessful)
                {
                    await this.retryPolicy.ExecuteAsync(async () =>
                    {
                        try
                        {
                            postEntity = await this.teamIdeaStorageProvider.GetPostAsync(postCreatedByUserId, postId);
                            postEntity.TotalVotes += 1;

                            // Update operation will throw exception if the column has already been updated
                            // or if there is a transient error (handled by an Azure storage)
                            await this.teamIdeaStorageProvider.UpsertTeamIdeaAsync(postEntity);
                            await this.teamIdeaSearchService.RunIndexerOnDemandAsync();
                        }
                        catch (StorageException ex)
                        {
                            if (ex.RequestInformation.HttpStatusCode == StatusCodes.Status412PreconditionFailed)
                            {
                                this.logger.LogError("Optimistic concurrency violation – entity has changed since it was retrieved.");
                                throw;
                            }
                        }
#pragma warning disable CA1031 // catching generic exception to trace log error in telemetry and continue the execution
                        catch (Exception ex)
#pragma warning restore CA1031 // catching generic exception to trace log error in telemetry and continue the execution
                        {
                            // log exception details to telemetry
                            // but do not attempt to retry in order to avoid multiple vote count decrement
                            this.logger.LogError(ex, "Exception occurred while reading post details.");
                        }
                    });
                }
                else
                {
                    this.logger.LogInformation($"User vote deleted for user{this.UserAadId} for PostId:{postId}");
                    await this.teamIdeaSearchService.RunIndexerOnDemandAsync();
                    return this.Ok(true);
                }
            }

            return this.Ok(false);
        }
    }
}