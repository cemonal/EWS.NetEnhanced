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
    /// Represents a time zone in which a meeting is defined.
    /// </summary>
    internal sealed class MeetingTimeZone : ComplexProperty
    {
        private string name;
        private TimeSpan? baseOffset;
        private TimeChange standard;
        private TimeChange daylight;

        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingTimeZone"/> class.
        /// </summary>
        /// <param name="timeZone">The time zone used to initialize this instance.</param>
        internal MeetingTimeZone(TimeZoneInfo timeZone)
        {
            // Unfortunately, MeetingTimeZone does not support all the time transition types
            // supported by TimeZoneInfo. That leaves us unable to accurately convert TimeZoneInfo
            // into MeetingTimeZone. So we don't... Instead, we emit the time zone's Id and
            // hope the server will find a match (which it should).
            this.Name = timeZone.Id;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingTimeZone"/> class.
        /// </summary>
        public MeetingTimeZone()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingTimeZone"/> class.
        /// </summary>
        /// <param name="name">The name of the time zone.</param>
        public MeetingTimeZone(string name)
            : this()
        {
            this.name = name;
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
                case XmlElementNames.BaseOffset:
                    this.baseOffset = EwsUtilities.XSDurationToTimeSpan(reader.ReadElementValue());
                    return true;
                case XmlElementNames.Standard:
                    this.standard = new TimeChange();
                    this.standard.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.Daylight:
                    this.daylight = new TimeChange();
                    this.daylight.LoadFromXml(reader, reader.LocalName);
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Reads the attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
        {
            this.name = reader.ReadAttributeValue(XmlAttributeNames.TimeZoneName);
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.TimeZoneName, this.Name);
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (this.BaseOffset.HasValue)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.BaseOffset,
                    EwsUtilities.TimeSpanToXSDuration(this.BaseOffset.Value));
            }

            if (this.Standard != null)
            {
                this.Standard.WriteToXml(writer, XmlElementNames.Standard);
            }

            if (this.Daylight != null)
            {
                this.Daylight.WriteToXml(writer, XmlElementNames.Daylight);
            }
        }

        /// <summary>
        /// Converts this meeting time zone into a TimeZoneInfo structure.
        /// </summary>
        /// <returns></returns>
        internal TimeZoneInfo ToTimeZoneInfo()
        {
            // MeetingTimeZone.ToTimeZoneInfo throws ArgumentNullException if name is null
            // TimeZoneName is optional, may not show in the response.
            if (string.IsNullOrEmpty(this.Name))
            {
                return null;
            }

            TimeZoneInfo result = null;

            try
            {
                result = TimeZoneInfo.FindSystemTimeZoneById(this.Name);
            }
            catch (TimeZoneNotFoundException)
            {
                // Could not find a time zone with that Id on the local system.
            }

            // Again, we cannot accurately convert MeetingTimeZone into TimeZoneInfo
            // because TimeZoneInfo doesn't support absolute date transitions. So if
            // there is no system time zone that has a matching Id, we return null.
            return result;
        }

        /// <summary>
        /// Gets or sets the name of the time zone.
        /// </summary>
        public string Name
        {
            get { return this.name; }
            set { this.SetFieldValue(ref this.name, value); }
        }

        /// <summary>
        /// Gets or sets the base offset of the time zone from the UTC time zone.
        /// </summary>
        public TimeSpan? BaseOffset
        {
            get { return this.baseOffset; }
            set { this.SetFieldValue(ref this.baseOffset, value); }
        }

        /// <summary>
        /// Gets or sets a TimeChange defining when the time changes to Standard Time.
        /// </summary>
        public TimeChange Standard
        {
            get { return this.standard; }
            set { this.SetFieldValue(ref this.standard, value); }
        }

        /// <summary>
        /// Gets or sets a TimeChange defining when the time changes to Daylight Saving Time.
        /// </summary>
        public TimeChange Daylight
        {
            get { return this.daylight; }
            set { this.SetFieldValue(ref this.daylight, value); }
        }
    }
}