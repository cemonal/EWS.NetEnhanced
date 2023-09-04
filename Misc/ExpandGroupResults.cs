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
    using System.Collections.Generic;
    using System.Collections.ObjectModel;

    /// <summary>
    /// Represents the results of an ExpandGroup operation.
    /// </summary>
    public sealed class ExpandGroupResults : IEnumerable<EmailAddress>
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="ExpandGroupResults"/> class.
        /// </summary>
        internal ExpandGroupResults()
        {
        }

        /// <summary>
        /// Gets the number of members that were returned by the ExpandGroup operation. Count might be
        /// less than the total number of members in the group, in which case the value of the
        /// IncludesAllMembers is false.
        /// </summary>
        public int Count => this.Members.Count;

        /// <summary>
        /// Gets a value indicating whether all the members of the group have been returned by ExpandGroup.
        /// </summary>
        public bool IncludesAllMembers { get; private set; }

        /// <summary>
        /// Gets the members of the expanded group.
        /// </summary>
        public Collection<EmailAddress> Members { get; } = new Collection<EmailAddress>();

        #region IEnumerable<EmailAddress> Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<EmailAddress> GetEnumerator()
        {
            return this.Members.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.Members.GetEnumerator();
        }

        #endregion

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsServiceXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.DLExpansion);
            if (!reader.IsEmptyElement)
            {
                int totalItemsInView = reader.ReadAttributeValue<int>(XmlAttributeNames.TotalItemsInView);
                this.IncludesAllMembers = reader.ReadAttributeValue<bool>(XmlAttributeNames.IncludesLastItemInRange);

                for (int i = 0; i < totalItemsInView; i++)
                {
                    EmailAddress emailAddress = new EmailAddress();

                    reader.ReadStartElement(XmlNamespace.Types, XmlElementNames.Mailbox);
                    emailAddress.LoadFromXml(reader, XmlElementNames.Mailbox);

                    this.Members.Add(emailAddress);
                }

                reader.ReadEndElement(XmlNamespace.Messages, XmlElementNames.DLExpansion);
            }
        }
    }
}