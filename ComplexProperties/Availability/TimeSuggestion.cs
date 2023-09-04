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
    using System.Collections.ObjectModel;

    /// <summary>
    /// Represents an availability time suggestion.
    /// </summary>
    public sealed class TimeSuggestion : ComplexProperty
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="TimeSuggestion"/> class.
        /// </summary>
        internal TimeSuggestion() : base() { }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if appropriate element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.MeetingTime:
                    this.MeetingTime = reader.ReadElementValueAsUnbiasedDateTimeScopedToServiceTimeZone();
                    return true;
                case XmlElementNames.IsWorkTime:
                    this.IsWorkTime = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.SuggestionQuality:
                    this.Quality = reader.ReadElementValue<SuggestionQuality>();
                    return true;
                case XmlElementNames.AttendeeConflictDataArray:
                    if (!reader.IsEmptyElement)
                    {
                        do
                        {
                            reader.Read();

                            if (!reader.IsStartElement())
                                continue;
                            
                            Conflict conflict = null;

                            switch (reader.LocalName)
                            {
                                case XmlElementNames.UnknownAttendeeConflictData:
                                    conflict = new Conflict(ConflictType.UnknownAttendeeConflict);
                                    break;
                                case XmlElementNames.TooBigGroupAttendeeConflictData:
                                    conflict = new Conflict(ConflictType.GroupTooBigConflict);
                                    break;
                                case XmlElementNames.IndividualAttendeeConflictData:
                                    conflict = new Conflict(ConflictType.IndividualAttendeeConflict);
                                    break;
                                case XmlElementNames.GroupAttendeeConflictData:
                                    conflict = new Conflict(ConflictType.GroupConflict);
                                    break;
                                default:
                                    EwsUtilities.Assert(
                                        false,
                                        "TimeSuggestion.TryReadElementFromXml",
                                        string.Format("The {0} element name does not map to any AttendeeConflict descendant.", reader.LocalName));

                                    // The following line to please the compiler
                                    break;
                            }

                            if (conflict == null) continue;

                            conflict.LoadFromXml(reader, reader.LocalName);
                            this.Conflicts.Add(conflict);
                        }

                        while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.AttendeeConflictDataArray));
                    }

                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Gets the suggested time.
        /// </summary>
        public DateTime MeetingTime { get; private set; }

        /// <summary>
        /// Gets a value indicating whether the suggested time is within working hours.
        /// </summary>
        public bool IsWorkTime { get; private set; }

        /// <summary>
        /// Gets the quality of the suggestion.
        /// </summary>
        public SuggestionQuality Quality { get; private set; }

        /// <summary>
        /// Gets a collection of conflicts at the suggested time.
        /// </summary>
        public Collection<Conflict> Conflicts { get; } = new Collection<Conflict>();
    }
}