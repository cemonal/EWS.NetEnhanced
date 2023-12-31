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
    using System.Collections.ObjectModel;

    /// <summary>
    /// Represents the set of actions available for a rule.
    /// </summary>
    public sealed class RuleActions : ComplexProperty
    {
        /// <summary>
        /// SMS recipient address type.
        /// </summary>
        private const string MobileType = "MOBILE";

        /// <summary>
        /// The CopyToFolder action.
        /// </summary>
        private FolderId copyToFolder;

        /// <summary>
        /// The Delete action.
        /// </summary>
        private bool delete;

        /// <summary>
        /// The MarkImportance action.
        /// </summary>
        private Importance? markImportance;

        /// <summary>
        /// The MarkAsRead action.
        /// </summary>
        private bool markAsRead;

        /// <summary>
        /// The MoveToFolder action.
        /// </summary>
        private FolderId moveToFolder;

        /// <summary>
        /// The PermanentDelete action.
        /// </summary>
        private bool permanentDelete;

        /// <summary>
        /// The ServerReplyWithMessage action.
        /// </summary>
        private ItemId serverReplyWithMessage;

        /// <summary>
        /// The StopProcessingRules action.
        /// </summary>
        private bool stopProcessingRules;

        /// <summary>
        /// Initializes a new instance of the <see cref="RulePredicates"/> class.
        /// </summary>
        internal RuleActions() : base()
        {
            this.AssignCategories = new StringList();
            this.ForwardAsAttachmentToRecipients = new EmailAddressCollection(XmlElementNames.Address);
            this.ForwardToRecipients = new EmailAddressCollection(XmlElementNames.Address);
            this.RedirectToRecipients = new EmailAddressCollection(XmlElementNames.Address);
            this.SendSMSAlertToRecipients = new Collection<MobilePhone>();
        }

        /// <summary>
        /// Gets the categories that should be stamped on incoming messages. 
        /// To disable stamping incoming messages with categories, set 
        /// AssignCategories to null.
        /// </summary>
        public StringList AssignCategories { get; }

        /// <summary>
        /// Gets or sets the Id of the folder incoming messages should be copied to.
        /// To disable copying incoming messages to a folder, set CopyToFolder to null.
        /// </summary>
        public FolderId CopyToFolder
        {
            get
            {
                return this.copyToFolder;
            }

            set
            {
                this.SetFieldValue(ref this.copyToFolder, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages should be
        /// automatically moved to the Deleted Items folder.
        /// </summary>
        public bool Delete
        {
            get
            {
                return this.delete;
            }

            set
            {
                this.SetFieldValue(ref this.delete, value);
            }
        }

        /// <summary>
        /// Gets the e-mail addresses to which incoming messages should be 
        /// forwarded as attachments. To disable forwarding incoming messages
        /// as attachments, empty the ForwardAsAttachmentToRecipients list.
        /// </summary>
        public EmailAddressCollection ForwardAsAttachmentToRecipients { get; }

        /// <summary>
        /// Gets the e-mail addresses to which incoming messages should be forwarded. 
        /// To disable forwarding incoming messages, empty the ForwardToRecipients list.
        /// </summary>
        public EmailAddressCollection ForwardToRecipients { get; }

        /// <summary>
        /// Gets or sets the importance that should be stamped on incoming 
        /// messages. To disable the stamping of incoming messages with an 
        /// importance, set MarkImportance to null.
        /// </summary>
        public Importance? MarkImportance
        {
            get
            {
                return this.markImportance;
            }

            set
            {
                this.SetFieldValue(ref this.markImportance, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages should be 
        /// marked as read.
        /// </summary>
        public bool MarkAsRead
        {
            get
            {
                return this.markAsRead;
            }

            set
            {
                this.SetFieldValue(ref this.markAsRead, value);
            }
        }

        /// <summary>
        /// Gets or sets the Id of the folder to which incoming messages should be
        /// moved. To disable the moving of incoming messages to a folder, set
        /// CopyToFolder to null.
        /// </summary>
        public FolderId MoveToFolder
        {
            get
            {
                return this.moveToFolder;
            }

            set
            {
                this.SetFieldValue(ref this.moveToFolder, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages should be 
        /// permanently deleted. When a message is permanently deleted, it is never 
        /// saved into the recipient's mailbox. To delete a message after it has 
        /// been saved into the recipient's mailbox, use the Delete action.
        /// </summary>
        public bool PermanentDelete
        {
            get
            {
                return this.permanentDelete;
            }

            set
            {
                this.SetFieldValue(ref this.permanentDelete, value);
            }
        }

        /// <summary>
        /// Gets the e-mail addresses to which incoming messages should be 
        /// redirecteded. To disable redirection of incoming messages, empty
        /// the RedirectToRecipients list. Unlike forwarded mail, redirected mail
        /// maintains the original sender and recipients. 
        /// </summary>
        public EmailAddressCollection RedirectToRecipients { get; }

        /// <summary>
        /// Gets the phone numbers to which an SMS alert should be sent. To disable
        /// sending SMS alerts for incoming messages, empty the 
        /// SendSMSAlertToRecipients list.
        /// </summary>
        public Collection<MobilePhone> SendSMSAlertToRecipients { get; private set; }

        /// <summary>
        /// Gets or sets the Id of the template message that should be sent
        /// as a reply to incoming messages. To disable automatic replies, set 
        /// ServerReplyWithMessage to null. 
        /// </summary>
        public ItemId ServerReplyWithMessage
        {
            get
            {
                return this.serverReplyWithMessage;
            }

            set
            {
                this.SetFieldValue(ref this.serverReplyWithMessage, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether subsequent rules should be
        /// evaluated. 
        /// </summary>
        public bool StopProcessingRules
        {
            get
            {
                return this.stopProcessingRules;
            }

            set
            {
                this.SetFieldValue(ref this.stopProcessingRules, value);
            }
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.AssignCategories:
                    this.AssignCategories.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.CopyToFolder:
                    reader.ReadStartElement(XmlNamespace.NotSpecified, XmlElementNames.FolderId);
                    this.copyToFolder = new FolderId();
                    this.copyToFolder.LoadFromXml(reader, XmlElementNames.FolderId);
                    reader.ReadEndElement(XmlNamespace.NotSpecified, XmlElementNames.CopyToFolder);
                    return true;
                case XmlElementNames.Delete:
                    this.delete = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.ForwardAsAttachmentToRecipients:
                    this.ForwardAsAttachmentToRecipients.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.ForwardToRecipients:
                    this.ForwardToRecipients.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.MarkImportance:
                    this.markImportance = reader.ReadElementValue<Importance>();
                    return true;
                case XmlElementNames.MarkAsRead:
                    this.markAsRead = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.MoveToFolder:
                    reader.ReadStartElement(XmlNamespace.NotSpecified, XmlElementNames.FolderId);
                    this.moveToFolder = new FolderId();
                    this.moveToFolder.LoadFromXml(reader, XmlElementNames.FolderId);
                    reader.ReadEndElement(XmlNamespace.NotSpecified, XmlElementNames.MoveToFolder);
                    return true;
                case XmlElementNames.PermanentDelete:
                    this.permanentDelete = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.RedirectToRecipients:
                    this.RedirectToRecipients.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.SendSMSAlertToRecipients:
                    EmailAddressCollection smsRecipientCollection = new EmailAddressCollection(XmlElementNames.Address);
                    smsRecipientCollection.LoadFromXml(reader, reader.LocalName);
                    this.SendSMSAlertToRecipients = ConvertSMSRecipientsFromEmailAddressCollectionToMobilePhoneCollection(smsRecipientCollection);
                    return true;
                case XmlElementNames.ServerReplyWithMessage:
                    this.serverReplyWithMessage = new ItemId();
                    this.serverReplyWithMessage.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.StopProcessingRules:
                    this.stopProcessingRules = reader.ReadElementValue<bool>();
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
            if (this.AssignCategories.Count > 0)
            {
                this.AssignCategories.WriteToXml(writer, XmlElementNames.AssignCategories);
            }

            if (this.CopyToFolder != null)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.CopyToFolder);
                this.CopyToFolder.WriteToXml(writer);
                writer.WriteEndElement();
            }

            if (this.Delete)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.Delete,
                    this.Delete);
            }

            if (this.ForwardAsAttachmentToRecipients.Count > 0)
            {
                this.ForwardAsAttachmentToRecipients.WriteToXml(writer, XmlElementNames.ForwardAsAttachmentToRecipients);
            }

            if (this.ForwardToRecipients.Count > 0)
            {
                this.ForwardToRecipients.WriteToXml(writer, XmlElementNames.ForwardToRecipients);
            }

            if (this.MarkImportance.HasValue)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.MarkImportance,
                    this.MarkImportance.Value);
            }

            if (this.MarkAsRead)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.MarkAsRead,
                    this.MarkAsRead);
            }

            if (this.MoveToFolder != null)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.MoveToFolder);
                this.MoveToFolder.WriteToXml(writer);
                writer.WriteEndElement();
            }

            if (this.PermanentDelete)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.PermanentDelete,
                    this.PermanentDelete);
            }

            if (this.RedirectToRecipients.Count > 0)
            {
                this.RedirectToRecipients.WriteToXml(writer, XmlElementNames.RedirectToRecipients);
            }

            if (this.SendSMSAlertToRecipients.Count > 0)
            {
                EmailAddressCollection emailCollection = ConvertSMSRecipientsFromMobilePhoneCollectionToEmailAddressCollection(this.SendSMSAlertToRecipients);
                emailCollection.WriteToXml(writer, XmlElementNames.SendSMSAlertToRecipients);
            }

            if (this.ServerReplyWithMessage != null)
            {
                this.ServerReplyWithMessage.WriteToXml(writer, XmlElementNames.ServerReplyWithMessage);
            }

            if (this.StopProcessingRules)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.StopProcessingRules,
                    this.StopProcessingRules);
            }
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void InternalValidate()
        {
            base.InternalValidate();
            EwsUtilities.ValidateParam(this.ForwardAsAttachmentToRecipients, "ForwardAsAttachmentToRecipients");
            EwsUtilities.ValidateParam(this.ForwardToRecipients, "ForwardToRecipients");
            EwsUtilities.ValidateParam(this.RedirectToRecipients, "RedirectToRecipients");
            foreach (MobilePhone sendSMSAlertToRecipient in this.SendSMSAlertToRecipients)
            {
                EwsUtilities.ValidateParam(sendSMSAlertToRecipient, "SendSMSAlertToRecipient");
            }
        }

        /// <summary>
        /// Convert the SMS recipient list from EmailAddressCollection type to MobilePhone collection type.
        /// </summary>
        /// <param name="emailCollection">Recipient list in EmailAddressCollection type.</param>
        /// <returns>A MobilePhone collection object containing all SMS recipient in MobilePhone type. </returns>
        private static Collection<MobilePhone> ConvertSMSRecipientsFromEmailAddressCollectionToMobilePhoneCollection(EmailAddressCollection emailCollection)
        {
            Collection<MobilePhone> mobilePhoneCollection = new Collection<MobilePhone>();
            foreach (EmailAddress emailAddress in emailCollection)
            {
                mobilePhoneCollection.Add(new MobilePhone(emailAddress.Name, emailAddress.Address));
            }

            return mobilePhoneCollection;
        }

        /// <summary>
        /// Convert the SMS recipient list from MobilePhone collection type to EmailAddressCollection type.
        /// </summary>
        /// <param name="recipientCollection">Recipient list in a MobilePhone collection type.</param>
        /// <returns>An EmailAddressCollection object containing recipients with "MOBILE" address type. </returns>
        private static EmailAddressCollection ConvertSMSRecipientsFromMobilePhoneCollectionToEmailAddressCollection(Collection<MobilePhone> recipientCollection)
        {
            EmailAddressCollection emailCollection = new EmailAddressCollection(XmlElementNames.Address);
            foreach (MobilePhone recipient in recipientCollection)
            {
                EmailAddress emailAddress = new EmailAddress(
                    recipient.Name,
                    recipient.PhoneNumber,
                    MobileType);
                emailCollection.Add(emailAddress);
            }

            return emailCollection;
        }
    }
}