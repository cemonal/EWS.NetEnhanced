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
    using System.Collections.Generic;

    /// <summary>
    /// Represents an UpdateItem request.
    /// </summary>
    internal sealed class UpdateItemRequest : MultiResponseServiceRequest<UpdateItemResponse>
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="UpdateItemRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal UpdateItemRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Gets a value indicating whether the TimeZoneContext SOAP header should be emitted.
        /// </summary>
        /// <value>
        ///     <c>true</c> if the time zone should be emitted; otherwise, <c>false</c>.
        /// </value>
        internal override bool EmitTimeZoneHeader
        {
            get
            {
                foreach (Item item in this.Items)
                {
                    if (item.GetIsTimeZoneHeaderRequired(true /* isUpdateOperation */))
                    {
                        return true;
                    }
                }

                return false;
            }
        }

        /// <summary>
        /// Validates the request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParamCollection(this.Items, "Items");
            for (int i = 0; i < this.Items.Count; i++)
            {
                if ((this.Items[i] == null) || this.Items[i].IsNew)
                {
                    throw new ArgumentException(string.Format(Strings.ItemToUpdateCannotBeNullOrNew, i));
                }
            }

            if (this.SavedItemsDestinationFolder != null)
            {
                this.SavedItemsDestinationFolder.Validate(this.Service.RequestedServerVersion);
            }

            // Validate each item.
            foreach (Item item in this.Items)
            {
                item.Validate();
            }

            if (this.SuppressReadReceipts && this.Service.RequestedServerVersion < ExchangeVersion.Exchange2013)
            {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.ParameterIncompatibleWithRequestVersion,
                        "SuppressReadReceipts",
                        ExchangeVersion.Exchange2013));
            }
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Response object.</returns>
        internal override UpdateItemResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new UpdateItemResponse(this.Items[responseIndex]);
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.UpdateItem;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>Xml element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.UpdateItemResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>Xml element name.</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.UpdateItemResponseMessage;
        }

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of items in response.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return this.Items.Count;
        }

        /// <summary>
        /// Writes XML attributes.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            base.WriteAttributesToXml(writer);

            if (this.MessageDisposition.HasValue)
            {
                writer.WriteAttributeValue(XmlAttributeNames.MessageDisposition, this.MessageDisposition);
            }

            if (this.SuppressReadReceipts)
            {
                writer.WriteAttributeValue(XmlAttributeNames.SuppressReadReceipts, true);
            }

            writer.WriteAttributeValue(XmlAttributeNames.ConflictResolution, this.ConflictResolutionMode);

            if (this.SendInvitationsOrCancellationsMode.HasValue)
            {
                writer.WriteAttributeValue(
                    XmlAttributeNames.SendMeetingInvitationsOrCancellations,
                    this.SendInvitationsOrCancellationsMode.Value);
            }
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (this.SavedItemsDestinationFolder != null)
            {
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.SavedItemFolderId);
                this.SavedItemsDestinationFolder.WriteToXml(writer);
                writer.WriteEndElement();
            }

            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.ItemChanges);

            foreach (Item item in this.Items)
            {
                item.WriteToXmlForUpdate(writer);
            }

            writer.WriteEndElement();
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2007_SP1;
        }

        /// <summary>
        /// Gets or sets the message disposition.
        /// </summary>
        /// <value>The message disposition.</value>
        public MessageDisposition? MessageDisposition { get; set; }

        /// <summary>
        /// Gets or sets the conflict resolution mode.
        /// </summary>
        /// <value>The conflict resolution mode.</value>
        public ConflictResolutionMode ConflictResolutionMode { get; set; }

        /// <summary>
        /// Gets or sets the send invitations or cancellations mode.
        /// </summary>
        /// <value>The send invitations or cancellations mode.</value>
        public SendInvitationsOrCancellationsMode? SendInvitationsOrCancellationsMode { get; set; }

        /// <summary>
        /// Gets or sets whether to suppress read receipts
        /// </summary>
        /// <value>Whether to suppress read receipts</value>
        public bool SuppressReadReceipts
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the items.
        /// </summary>
        /// <value>The items.</value>
        public List<Item> Items { get; } = new List<Item>();

        /// <summary>
        /// Gets or sets the saved items destination folder.
        /// </summary>
        /// <value>The saved items destination folder.</value>
        public FolderId SavedItemsDestinationFolder { get; set; }
    }
}