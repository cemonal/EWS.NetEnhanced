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
    /// Encapsulates information on the occurrence of a recurring appointment.
    /// </summary>
    public sealed class OccurrenceInfo : ComplexProperty
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="OccurrenceInfo"/> class.
        /// </summary>
        internal OccurrenceInfo()
        {
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
                case XmlElementNames.ItemId:
                    this.ItemId = new ItemId();
                    this.ItemId.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.Start:
                    this.Start = reader.ReadElementValueAsDateTime().Value;
                    return true;
                case XmlElementNames.End:
                    this.End = reader.ReadElementValueAsDateTime().Value;
                    return true;
                case XmlElementNames.OriginalStart:
                    this.OriginalStart = reader.ReadElementValueAsDateTime().Value;
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Gets the Id of the occurrence.
        /// </summary>
        public ItemId ItemId { get; private set; }

        /// <summary>
        /// Gets the start date and time of the occurrence.
        /// </summary>
        public DateTime Start { get; private set; }

        /// <summary>
        /// Gets the end date and time of the occurrence.
        /// </summary>
        public DateTime End { get; private set; }

        /// <summary>
        /// Gets the original start date and time of the occurrence.
        /// </summary>
        public DateTime OriginalStart { get; private set; }
    }
}