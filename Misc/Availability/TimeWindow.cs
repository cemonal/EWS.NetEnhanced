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
    using System.Globalization;

    /// <summary>
    /// Represents a time period.
    /// </summary>
    public sealed class TimeWindow : ISelfValidate
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="TimeWindow"/> class.
        /// </summary>
        internal TimeWindow()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TimeWindow"/> class.
        /// </summary>
        /// <param name="startTime">The start date and time.</param>
        /// <param name="endTime">The end date and time.</param>
        public TimeWindow(DateTime startTime, DateTime endTime)
            : this()
        {
            this.StartTime = startTime;
            this.EndTime = endTime;
        }

        /// <summary>
        /// Gets or sets the start date and time.
        /// </summary>
        public DateTime StartTime { get; set; }

        /// <summary>
        /// Gets or sets the end date and time.
        /// </summary>
        public DateTime EndTime { get; set; }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsServiceXmlReader reader)
        {
            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, XmlElementNames.Duration);

            this.StartTime = reader.ReadElementValueAsDateTime(XmlNamespace.Types, XmlElementNames.StartTime).Value;
            this.EndTime = reader.ReadElementValueAsDateTime(XmlNamespace.Types, XmlElementNames.EndTime).Value;

            reader.ReadEndElement(XmlNamespace.Types, XmlElementNames.Duration);
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="startTime">The start time.</param>
        /// <param name="endTime">The end time.</param>
        private static void WriteToXml(
            EwsServiceXmlWriter writer,
            string xmlElementName,
            object startTime,
            object endTime)
        {
            writer.WriteStartElement(XmlNamespace.Types, xmlElementName);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.StartTime,
                startTime);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.EndTime,
                endTime);

            writer.WriteEndElement(); // xmlElementName
        }

        /// <summary>
        /// Writes to XML without scoping the dates and without emitting times.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal void WriteToXmlUnscopedDatesOnly(EwsServiceXmlWriter writer, string xmlElementName)
        {
            const string DateOnlyFormat = "yyyy-MM-ddT00:00:00";

            WriteToXml(
                writer,
                xmlElementName,
                this.StartTime.ToString(DateOnlyFormat, CultureInfo.InvariantCulture),
                this.EndTime.ToString(DateOnlyFormat, CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer, string xmlElementName)
        {
            WriteToXml(
                writer,
                xmlElementName,
                this.StartTime,
                this.EndTime);
        }

        /// <summary>
        /// Gets the duration.
        /// </summary>
        internal TimeSpan Duration => this.EndTime - this.StartTime;

        #region ISelfValidate Members

        /// <summary>
        /// Validates this instance.
        /// </summary>
        void ISelfValidate.Validate()
        {
            if (this.StartTime >= this.EndTime)
            {
                throw new ArgumentException(Strings.TimeWindowStartTimeMustBeGreaterThanEndTime);
            }
        }

        #endregion
    }
}