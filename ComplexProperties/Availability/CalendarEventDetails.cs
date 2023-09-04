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
    /// Represents the details of a calendar event as returned by the GetUserAvailability operation.
    /// </summary>
    public sealed class CalendarEventDetails : ComplexProperty
    {
        private bool isPrivate;

        /// <summary>
        /// Initializes a new instance of the <see cref="CalendarEventDetails"/> class.
        /// </summary>
        internal CalendarEventDetails()
            : base()
        {
        }

        /// <summary>
        /// Attempts to read the element at the reader's current position.
        /// </summary>
        /// <param name="reader">The reader used to read the element.</param>
        /// <returns>True if the element was read, false otherwise.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.ID:
                    this.StoreId = reader.ReadElementValue();
                    return true;
                case XmlElementNames.Subject:
                    this.Subject = reader.ReadElementValue();
                    return true;
                case XmlElementNames.Location:
                    this.Location = reader.ReadElementValue();
                    return true;
                case XmlElementNames.IsMeeting:
                    this.IsMeeting = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsRecurring:
                    this.IsRecurring = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsException:
                    this.IsException = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsReminderSet:
                    this.IsReminderSet = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsPrivate:
                    this.isPrivate = reader.ReadElementValue<bool>();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Gets the store Id of the calendar event.
        /// </summary>
        public string StoreId { get; private set; }

        /// <summary>
        /// Gets the subject of the calendar event.
        /// </summary>
        public string Subject { get; private set; }

        /// <summary>
        /// Gets the location of the calendar event.
        /// </summary>
        public string Location { get; private set; }

        /// <summary>
        /// Gets a value indicating whether the calendar event is a meeting.
        /// </summary>
        public bool IsMeeting { get; private set; }

        /// <summary>
        /// Gets a value indicating whether the calendar event is recurring.
        /// </summary>
        public bool IsRecurring { get; private set; }

        /// <summary>
        /// Gets a value indicating whether the calendar event is an exception in a recurring series.
        /// </summary>
        public bool IsException { get; private set; }

        /// <summary>
        /// Gets a value indicating whether the calendar event has a reminder set.
        /// </summary>
        public bool IsReminderSet { get; private set; }

        /// <summary>
        /// Gets a value indicating whether the calendar event is private.
        /// </summary>
        public bool IsPrivate => this.isPrivate;
    }
}