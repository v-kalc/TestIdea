// <copyright file="TagsValidationAttribute.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.SubmitIdea.Helpers.CustomValidations
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.Globalization;

    /// <summary>
    /// Validate tag based on length and tag count for post.
    /// </summary>
    public sealed class TagsValidationAttribute : ValidationAttribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TagsValidationAttribute"/> class.
        /// </summary>
        /// <param name="tagsMaxCount">Max count of tags for validation.</param>
        public TagsValidationAttribute(int tagsMaxCount)
        {
            this.TagsMaxCount = tagsMaxCount;
        }

        /// <summary>
        /// Gets max count of tags for validation.
        /// </summary>
        public int TagsMaxCount { get; }

        /// <summary>
        /// Validate tag based on tag length and number of tags separated by comma.
        /// </summary>
        /// <param name="value">String containing tags separated by comma.</param>
        /// <param name="validationContext">Context for getting object which needs to be validated.</param>
        /// <returns>Validation result (either error message for failed validation or success).</returns>
        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {
            var tags = Convert.ToString(value, CultureInfo.InvariantCulture);

            if (string.IsNullOrEmpty(tags))
            {
                var tagsList = tags.Split(';');

                if (tagsList.Length > this.TagsMaxCount)
                {
                    return new ValidationResult("Max tags count exceeded");
                }

                foreach (var tag in tagsList)
                {
                    if (string.IsNullOrWhiteSpace(tag))
                    {
                        return new ValidationResult("Tag cannot be null or empty");
                    }

                    if (tag.Length > 20)
                    {
                        return new ValidationResult("Max tag length exceeded");
                    }
                }
            }

            // Tags are not mandatory for adding/updating post
            return ValidationResult.Success;
        }
    }
}
