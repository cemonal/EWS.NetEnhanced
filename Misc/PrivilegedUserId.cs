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

    /// <summary>
    /// Represents an privileged user Id.
    /// </summary>
    internal sealed class PrivilegedUserId
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="PrivilegedUserId"/> class.
        /// </summary>
        public PrivilegedUserId()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PrivilegedUserId"/> class.
        /// </summary>
        /// <param name="openType">The open type.</param>
        /// <param name="idType">The type of this Id.</param>
        /// <param name="id">The user Id.</param>
        public PrivilegedUserId(PrivilegedLogonType openType, ConnectingIdType idType, string id)
            : this()
        {
            this.LogonType = openType;
            this.IdType = idType;
            this.Id = id;
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="requestedServerVersion">The requested server version.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer, ExchangeVersion requestedServerVersion)
        {
            if (string.IsNullOrEmpty(this.Id))
            {
                throw new ArgumentException(Strings.IdPropertyMustBeSet);
            }

            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.OpenAsAdminOrSystemService);
            writer.WriteAttributeString(XmlElementNames.LogonType, this.LogonType.ToString());
            if (requestedServerVersion >= ExchangeVersion.Exchange2013 && this.BudgetType.HasValue)
            {
                writer.WriteAttributeString(XmlElementNames.BudgetType, ((int)this.BudgetType.Value).ToString());
            }

            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.ConnectingSID);
            writer.WriteElementValue(XmlNamespace.Types, this.IdType.ToString(), this.Id);
            writer.WriteEndElement(); // ConnectingSID
            writer.WriteEndElement(); // OpenAsAdminOrSystemService
        }

        /// <summary>
        /// Gets or sets the type of the Id.
        /// </summary>
        public ConnectingIdType IdType { get; set; }

        /// <summary>
        /// Gets or sets the user Id.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets the special logon type.
        /// </summary>
        public PrivilegedLogonType LogonType { get; set; }

        /// <summary>
        /// Gets or sets the budget type.
        /// </summary>
        public PrivilegedUserIdBudgetType? BudgetType { get; set; }
    }

    /// <summary>
    /// PrivilegedUserId BudgetType enum
    /// </summary>
    internal enum PrivilegedUserIdBudgetType
    {
        /// <summary>
        /// Interactive, charge against a copy of target mailbox budget.
        /// </summary>
        Default = 0,

        /// <summary>
        /// Running as background load
        /// </summary>
        RunningAsBackgroundLoad = 1,

        /// <summary>
        /// Unthrottled budget.
        /// </summary>
        Unthrottled = 2,
    }
}