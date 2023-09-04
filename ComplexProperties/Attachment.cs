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
    /// Represents an attachment to an item.
    /// </summary>
    public abstract class Attachment : ComplexProperty
    {
        private string name;
        private string contentType;
        private string contentId;
        private string contentLocation;
        private int size;
        private DateTime lastModifiedTime;
        private bool isInline;

        /// <summary>
        /// Initializes a new instance of the <see cref="Attachment"/> class.
        /// </summary>
        /// <param name="owner">The owner.</param>
        protected Attachment(Item owner)
        {
            this.Owner = owner;

            if (owner != null)
            {
                this.Service = this.Owner.Service;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Attachment"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        protected Attachment(ExchangeService service)
        {
            this.Service = service;
        }

        /// <summary>
        /// Throws exception if this is not a new service object.
        /// </summary>
        internal void ThrowIfThisIsNotNew()
        {
            if (!this.IsNew)
            {
                throw new InvalidOperationException(Strings.AttachmentCannotBeUpdated);
            }
        }

        /// <summary>
        /// Sets value of field.
        /// </summary>
        /// <remarks>
        /// We override the base implementation. Attachments cannot be modified so any attempts
        /// the change a property on an existing attachment is an error.
        /// </remarks>
        /// <typeparam name="T">Field type.</typeparam>
        /// <param name="field">The field.</param>
        /// <param name="value">The value.</param>
        internal override void SetFieldValue<T>(ref T field, T value)
        {
            this.ThrowIfThisIsNotNew();
            base.SetFieldValue(ref field, value);
        }

        /// <summary>
        /// Gets the Id of the attachment.
        /// </summary>
        public string Id { get; internal set; }

        /// <summary>
        /// Gets or sets the name of the attachment.
        /// </summary>
        public string Name
        {
            get { return this.name; }
            set { this.SetFieldValue(ref this.name, value); }
        }

        /// <summary>
        /// Gets or sets the content type of the attachment.
        /// </summary>
        public string ContentType
        {
            get { return this.contentType; }
            set { this.SetFieldValue(ref this.contentType, value); }
        }

        /// <summary>
        /// Gets or sets the content Id of the attachment. ContentId can be used as a custom way to identify
        /// an attachment in order to reference it from within the body of the item the attachment belongs to.
        /// </summary>
        public string ContentId
        {
            get { return this.contentId; }
            set { this.SetFieldValue(ref this.contentId, value); }
        }

        /// <summary>
        /// Gets or sets the content location of the attachment. ContentLocation can be used to associate
        /// an attachment with a Url defining its location on the Web.
        /// </summary>
        public string ContentLocation
        {
            get { return this.contentLocation; }
            set { this.SetFieldValue(ref this.contentLocation, value); }
        }

        /// <summary>
        /// Gets the size of the attachment.
        /// </summary>
        public int Size
        {
            get
            {
                EwsUtilities.ValidatePropertyVersion(this.Service, ExchangeVersion.Exchange2010, "Size");

                return this.size;
            }

            internal set
            {
                EwsUtilities.ValidatePropertyVersion(this.Service, ExchangeVersion.Exchange2010, "Size");

                this.SetFieldValue(ref this.size, value);
            }
        }

        /// <summary>
        /// Gets the date and time when this attachment was last modified.
        /// </summary>
        public DateTime LastModifiedTime
        {
            get
            {
                EwsUtilities.ValidatePropertyVersion(this.Service, ExchangeVersion.Exchange2010, "LastModifiedTime");

                return this.lastModifiedTime;
            }

            internal set
            {
                EwsUtilities.ValidatePropertyVersion(this.Service, ExchangeVersion.Exchange2010, "LastModifiedTime");

                this.SetFieldValue(ref this.lastModifiedTime, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this is an inline attachment.
        /// Inline attachments are not visible to end users.
        /// </summary>
        public bool IsInline
        {
            get
            {
                EwsUtilities.ValidatePropertyVersion(this.Service, ExchangeVersion.Exchange2010, "IsInline");

                return this.isInline;
            }

            set
            {
                EwsUtilities.ValidatePropertyVersion(this.Service, ExchangeVersion.Exchange2010, "IsInline");

                this.SetFieldValue(ref this.isInline, value);
            }
        }

        /// <summary>
        /// True if the attachment has not yet been saved, false otherwise.
        /// </summary>
        internal bool IsNew => string.IsNullOrEmpty(this.Id);

        /// <summary>
        /// Gets the owner of the attachment.
        /// </summary>
        internal Item Owner { get; }

        /// <summary>
        /// Gets the related exchange service.
        /// </summary>
        internal ExchangeService Service { get; }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal abstract string GetXmlElementName();

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.AttachmentId:
                    this.Id = reader.ReadAttributeValue(XmlAttributeNames.Id);

                    if (this.Owner != null)
                    {
                        string rootItemChangeKey = reader.ReadAttributeValue(XmlAttributeNames.RootItemChangeKey);

                        if (!string.IsNullOrEmpty(rootItemChangeKey))
                        {
                            this.Owner.RootItemId.ChangeKey = rootItemChangeKey;
                        }
                    }
                    reader.ReadEndElementIfNecessary(XmlNamespace.Types, XmlElementNames.AttachmentId);
                    return true;
                case XmlElementNames.Name:
                    this.name = reader.ReadElementValue();
                    return true;
                case XmlElementNames.ContentType:
                    this.contentType = reader.ReadElementValue();
                    return true;
                case XmlElementNames.ContentId:
                    this.contentId = reader.ReadElementValue();
                    return true;
                case XmlElementNames.ContentLocation:
                    this.contentLocation = reader.ReadElementValue();
                    return true;
                case XmlElementNames.Size:
                    this.size = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.LastModifiedTime:
                    this.lastModifiedTime = reader.ReadElementValueAsDateTime().Value;
                    return true;
                case XmlElementNames.IsInline:
                    this.isInline = reader.ReadElementValue<bool>();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Name, this.Name);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ContentType, this.ContentType);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ContentId, this.ContentId);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ContentLocation, this.ContentLocation);
            if (writer.Service.RequestedServerVersion > ExchangeVersion.Exchange2007_SP1)
            {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsInline, this.IsInline);
            }
        }

        /// <summary>
        /// Load the attachment.
        /// </summary>
        /// <param name="bodyType">Type of the body.</param>
        /// <param name="additionalProperties">The additional properties.</param>
        internal void InternalLoad(BodyType? bodyType, IEnumerable<PropertyDefinitionBase> additionalProperties)
        {
            this.Service.GetAttachment(
                this,
                bodyType,
                additionalProperties);
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        /// <param name="attachmentIndex">Index of this attachment.</param>
        internal virtual void Validate(int attachmentIndex)
        {
        }

        /// <summary>
        /// Loads the attachment. Calling this method results in a call to EWS.
        /// </summary>
        public void Load()
        {
            this.InternalLoad(null, null);
        }
    }
}