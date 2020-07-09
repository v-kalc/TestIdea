// <copyright file="BaseSubmitIdeaController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Error = Microsoft.Teams.Apps.SubmitIdea.Models.ErrorResponse;

    /// <summary>
    /// Base controller to handle ideas API operations.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class BaseSubmitIdeaController : ControllerBase
    {
        /// <summary>
        /// Instance of application insights telemetry client.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="BaseSubmitIdeaController"/> class.
        /// </summary>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        public BaseSubmitIdeaController(TelemetryClient telemetryClient)
        {
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Gets the user tenant id from the HttpContext.
        /// </summary>
        protected string UserTenantId
        {
            get
            {
                var tenantClaimType = "http://schemas.microsoft.com/identity/claims/tenantid";
                var claim = this.User.Claims.FirstOrDefault(p => tenantClaimType.Equals(p.Type, StringComparison.OrdinalIgnoreCase));
                if (claim == null)
                {
                    return null;
                }

                return claim.Value;
            }
        }

        /// <summary>
        /// Gets the user Azure Active Directory id from the HttpContext.
        /// </summary>
        protected string UserAadId
        {
            get
            {
                var oidClaimType = "http://schemas.microsoft.com/identity/claims/objectidentifier";
                var claim = this.User.Claims.FirstOrDefault(p => oidClaimType.Equals(p.Type, StringComparison.OrdinalIgnoreCase));
                if (claim == null)
                {
                    return null;
                }

                return claim.Value;
            }
        }

        /// <summary>
        /// Gets the user name from the HttpContext.
        /// </summary>
        protected string UserName
        {
            get
            {
                var claim = this.User.Claims.FirstOrDefault(p => "name".Equals(p.Type, StringComparison.OrdinalIgnoreCase));
                if (claim == null)
                {
                    return null;
                }

                return claim.Value;
            }
        }

        /// <summary>
        /// Records event data to Application Insights telemetry client.
        /// </summary>
        /// <param name="eventName">Name of the event.</param>
        public void RecordEvent(string eventName)
        {
            this.telemetryClient.TrackEvent(eventName, new Dictionary<string, string>
            {
                { "userId", this.UserAadId },
            });
        }

        /// <summary>
        /// Creates the error response as per the status codes.
        /// </summary>
        /// <param name="statusCode">Describes the type of error.</param>
        /// <param name="errorMessage">Describes the error message.</param>
        /// <returns>Returns error response with appropriate message and status code.</returns>
        protected IActionResult GetErrorResponse(int statusCode, string errorMessage)
        {
            switch (statusCode)
            {
                case StatusCodes.Status400BadRequest:
                    return this.StatusCode(
                      StatusCodes.Status400BadRequest,
                      new Error
                      {
                          StatusCode = "badRequest",
                          ErrorMessage = errorMessage,
                      });
                default:
                    return this.StatusCode(
                      StatusCodes.Status500InternalServerError,
                      new Error
                      {
                          StatusCode = "internalServerError",
                          ErrorMessage = errorMessage,
                      });
            }
        }
    }
}