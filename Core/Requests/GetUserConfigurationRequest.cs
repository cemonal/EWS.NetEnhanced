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
    /// <summary>
    /// Represents a GetUserConfiguration request.
    /// </summary>
    internal class GetUserConfigurationRequest : MultiResponseServiceRequest<GetUserConfigurationResponse>
    {
        private const string EnumDelimiter = ",";
        private UserConfiguration userConfiguration;
        
        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            EwsUtilities.ValidateParam(this.Name, "name");
            EwsUtilities.ValidateParam(this.ParentFolderId, "parentFolderId");
            this.ParentFolderId.Validate(this.Service.RequestedServerVersion);
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override GetUserConfigurationResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            // In the case of UserConfiguration.Load(), this.userConfiguration is set.
            if (this.userConfiguration == null)
            {
                this.userConfiguration = new UserConfiguration(service, this.Properties)
                {
                    Name = this.Name,
                    ParentFolderId = this.ParentFolderId
                };
            }

            return new GetUserConfigurationResponse(this.userConfiguration);
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2010;
        }

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected response messages.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return 1;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetUserConfiguration;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetUserConfigurationResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.GetUserConfigurationResponseMessage;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            // Write UserConfiguationName element
            UserConfiguration.WriteUserConfigurationNameToXml(
                writer,
                XmlNamespace.Messages,
                this.Name,
                this.ParentFolderId);

            // Write UserConfigurationProperties element
            writer.WriteElementValue(
                XmlNamespace.Messages,
                XmlElementNames.UserConfigurationProperties,
                this.Properties.ToString().Replace(EnumDelimiter, string.Empty));
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GetUserConfigurationRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal GetUserConfigurationRequest(ExchangeService service)
            : base(service, ServiceErrorHandling.ThrowOnError)
        {
        }

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>The name.</value>
        internal string Name { get; set; }

        /// <summary>
        /// Gets or sets the parent folder Id.
        /// </summary>
        /// <value>The parent folder Id.</value>
        internal FolderId ParentFolderId { get; set; }

        /// <summary>
        /// Gets or sets the user configuration.
        /// </summary>
        /// <value>The user configuration.</value>
        internal UserConfiguration UserConfiguration
        {
            get { return this.userConfiguration; }

            set 
            { 
                this.userConfiguration = value;

                this.Name = this.userConfiguration.Name;
                this.ParentFolderId = this.userConfiguration.ParentFolderId;
            }
        }

        /// <summary>
        /// Gets or sets the properties.
        /// </summary>
        /// <value>The properties.</value>
        internal UserConfigurationProperties Properties { get; set; }
    }
}