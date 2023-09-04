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
    /// Represents a conflict in a meeting time suggestion.
    /// </summary>
    public sealed class Conflict : ComplexProperty
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="Conflict"/> class.
        /// </summary>
        /// <param name="conflictType">The type of the conflict.</param>
        internal Conflict(ConflictType conflictType)
            : base()
        {
            this.ConflictType = conflictType;
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if appropriate element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.NumberOfMembers:
                    this.NumberOfMembers = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.NumberOfMembersAvailable:
                    this.NumberOfMembersAvailable = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.NumberOfMembersWithConflict:
                    this.NumberOfMembersWithConflict = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.NumberOfMembersWithNoData:
                    this.NumberOfMembersWithNoData = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.BusyType:
                    this.FreeBusyStatus = reader.ReadElementValue<LegacyFreeBusyStatus>();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Gets the type of the conflict.
        /// </summary>
        public ConflictType ConflictType { get; }

        /// <summary>
        /// Gets the number of users, resources, and rooms in the conflicting group. The value of this property
        /// is only meaningful when ConflictType is equal to ConflictType.GroupConflict.
        /// </summary>
        public int NumberOfMembers { get; private set; }

        /// <summary>
        /// Gets the number of members who are available (whose status is Free) in the conflicting group. The value
        /// of this property is only meaningful when ConflictType is equal to ConflictType.GroupConflict.
        /// </summary>
        public int NumberOfMembersAvailable { get; private set; }

        /// <summary>
        /// Gets the number of members who have a conflict (whose status is Busy, OOF or Tentative) in the conflicting
        /// group. The value of this property is only meaningful when ConflictType is equal to ConflictType.GroupConflict.
        /// </summary>
        public int NumberOfMembersWithConflict { get; private set; }

        /// <summary>
        /// Gets the number of members who do not have published free/busy data in the conflicting group. The value
        /// of this property is only meaningful when ConflictType is equal to ConflictType.GroupConflict.
        /// </summary>
        public int NumberOfMembersWithNoData { get; private set; }

        /// <summary>
        /// Gets the free/busy status of the conflicting attendee. The value of this property is only meaningful when
        /// ConflictType is equal to ConflictType.IndividualAttendee.
        /// </summary>
        public LegacyFreeBusyStatus FreeBusyStatus { get; private set; }
    }
}