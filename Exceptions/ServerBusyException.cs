/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace EWS.NetEnhanced.Data
{
    using System;
    using System.Runtime.Serialization;

    /// <summary>
    /// Represents a server busy exception found in a service response.
    /// </summary>
    [Serializable]
    public class ServerBusyException : ServiceResponseException
    {
        private const string BackOffMillisecondsKey = @"BackOffMilliseconds";
        private readonly int backOffMilliseconds;

        /// <summary>
        /// Initializes a new instance of the <see cref="ServerBusyException"/> class.
        /// </summary>
        /// <param name="response">The ServiceResponse when service operation failed remotely.</param>
        public ServerBusyException(ServiceResponse response) : base(response)
        {
            if (response.ErrorDetails != null && response.ErrorDetails.ContainsKey(BackOffMillisecondsKey))
            {
                int.TryParse(response.ErrorDetails[BackOffMillisecondsKey], out this.backOffMilliseconds);
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="T:EWS.NetEnhanced.Data.ServerBusyException"/> class with serialized data.
        /// </summary>
        /// <param name="info">The object that holds the serialized object data.</param>
        /// <param name="context">The contextual information about the source or destination.</param>
        protected ServerBusyException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
            this.backOffMilliseconds = info.GetInt32("BackOffMilliseconds");
        }

        /// <summary>Sets the <see cref="T:System.Runtime.Serialization.SerializationInfo" /> object with the parameter name and additional exception information.</summary>
        /// <param name="info">The object that holds the serialized object data. </param>
        /// <param name="context">The contextual information about the source or destination. </param>
        /// <exception cref="T:System.ArgumentNullException">The <paramref name="info" /> object is a null reference (Nothing in Visual Basic). </exception>
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            EwsUtilities.Assert(info != null, "ServerBusyException.GetObjectData", "info is null");

            base.GetObjectData(info, context);

            info?.AddValue("BackOffMilliseconds", this.backOffMilliseconds);
        }

        /// <summary>
        /// Suggested number of milliseconds to wait before attempting a request again. If zero, 
        /// there is no suggested backoff time.
        /// </summary>
        public int BackOffMilliseconds => this.backOffMilliseconds;
    }
}