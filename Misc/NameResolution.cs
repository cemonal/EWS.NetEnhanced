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
    /// Represents a suggested name resolution.
    /// </summary>
    public sealed class NameResolution
    {
        private readonly NameResolutionCollection owner;

        /// <summary>
        /// Initializes a new instance of the <see cref="NameResolution"/> class.
        /// </summary>
        /// <param name="owner">The owner.</param>
        internal NameResolution(NameResolutionCollection owner)
        {
            EwsUtilities.Assert(
                owner != null,
                "NameResolution.ctor",
                "owner is null.");

            this.owner = owner;
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsServiceXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Types, XmlElementNames.Resolution);

            reader.ReadStartElement(XmlNamespace.Types, XmlElementNames.Mailbox);
            this.Mailbox.LoadFromXml(reader, XmlElementNames.Mailbox);

            reader.Read();
            if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Contact))
            {
                this.Contact = new Contact(this.owner.Session);

                // Contacts returned by ResolveNames should behave like Contact.Load with FirstClassPropertySet specified.
                this.Contact.LoadFromXml(
                                    reader,
                                    true,                               /* clearPropertyBag */
                                    PropertySet.FirstClassProperties,
                                    false);                             /* summaryPropertiesOnly */

                reader.ReadEndElement(XmlNamespace.Types, XmlElementNames.Resolution);
            }
            else
            {
                reader.EnsureCurrentNodeIsEndElement(XmlNamespace.Types, XmlElementNames.Resolution);
            }
        }

        /// <summary>
        /// Gets the mailbox of the suggested resolved name.
        /// </summary>
        public EmailAddress Mailbox { get; } = new EmailAddress();

        /// <summary>
        /// Gets the contact information of the suggested resolved name. This property is only available when
        /// ResolveName is called with returnContactDetails = true.
        /// </summary>
        public Contact Contact { get; private set; }
    }
}