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
    /// Represents a custom time zone time change. 
    /// </summary>
    internal sealed class LegacyAvailabilityTimeZoneTime : ComplexProperty
    {
        private TimeSpan timeOfDay;

        /// <summary>
        /// Initializes a new instance of the <see cref="LegacyAvailabilityTimeZoneTime"/> class.
        /// </summary>
        internal LegacyAvailabilityTimeZoneTime()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="LegacyAvailabilityTimeZoneTime"/> class.
        /// </summary>
        /// <param name="transitionTime">The transition time used to initialize this instance.</param>
        /// <param name="delta">The offset used to initialize this instance.</param>
        internal LegacyAvailabilityTimeZoneTime(TimeZoneInfo.TransitionTime transitionTime, TimeSpan delta)
            : this()
        {
            this.Delta = delta;

            if (transitionTime.IsFixedDateRule)
            {
                // TimeZoneInfo doesn't support an actual year. Fixed date transitions occur at the same
                // date every year the adjustment rule the transition belongs to applies. The best thing
                // we can do here is use the current year.
                this.Year = DateTime.Today.Year;
                this.Month = transitionTime.Month;
                this.DayOrder = transitionTime.Day;
                this.timeOfDay = transitionTime.TimeOfDay.TimeOfDay;
            }
            else
            {
                // For floating rules, the mapping is direct.
                this.Year = 0;
                this.Month = transitionTime.Month;
                this.DayOfTheWeek = EwsUtilities.SystemToEwsDayOfTheWeek(transitionTime.DayOfWeek);
                this.DayOrder = transitionTime.Week;
                this.timeOfDay = transitionTime.TimeOfDay.TimeOfDay;
            }
        }

        /// <summary>
        /// Converts this instance to TimeZoneInfo.TransitionTime.
        /// </summary>
        /// <returns>A TimeZoneInfo.TransitionTime</returns>
        internal TimeZoneInfo.TransitionTime ToTransitionTime()
        {
            if (this.Year == 0)
            {
                return TimeZoneInfo.TransitionTime.CreateFloatingDateRule(
                    new DateTime(
                        DateTime.MinValue.Year,
                        DateTime.MinValue.Month,
                        DateTime.MinValue.Day,
                        this.timeOfDay.Hours,
                        this.timeOfDay.Minutes,
                        this.timeOfDay.Seconds),
                    this.Month,
                    this.DayOrder,
                    EwsUtilities.EwsToSystemDayOfWeek(this.DayOfTheWeek));
            }
            else
            {
                return TimeZoneInfo.TransitionTime.CreateFixedDateRule(
                    new DateTime(this.timeOfDay.Ticks),
                    this.Month,
                    this.DayOrder);
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
                case XmlElementNames.Bias:
                    this.Delta = TimeSpan.FromMinutes(reader.ReadElementValue<int>());
                    return true;
                case XmlElementNames.Time:
                    this.timeOfDay = TimeSpan.Parse(reader.ReadElementValue());
                    return true;
                case XmlElementNames.DayOrder:
                    this.DayOrder = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.DayOfWeek:
                    this.DayOfTheWeek = reader.ReadElementValue<DayOfTheWeek>();
                    return true;
                case XmlElementNames.Month:
                    this.Month = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.Year:
                    this.Year = reader.ReadElementValue<int>();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Writes the elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.Bias,
                (int)this.Delta.TotalMinutes);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.Time,
                EwsUtilities.TimeSpanToXSTime(this.timeOfDay));

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.DayOrder,
                this.DayOrder);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.Month,
                Month);

            // Only write DayOfWeek if this is a recurring time change
            if (this.Year == 0)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.DayOfWeek,
                    this.DayOfTheWeek);
            }

            // Only emit year if it's non zero, otherwise AS returns "Request is invalid"
            if (this.Year != 0)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.Year,
                    this.Year);
            }
        }

        /// <summary>
        /// Gets if current time presents DST transition time
        /// </summary>
        internal bool HasTransitionTime => this.Month >= 1 && this.Month <= 12;

        /// <summary>
        /// Gets or sets the delta.
        /// </summary>
        internal TimeSpan Delta { get; set; }

        /// <summary>
        /// Gets or sets the time of day.
        /// </summary>
        internal TimeSpan TimeOfDay
        {
            get { return this.timeOfDay; }
            set { this.timeOfDay = value; }
        }

        /// <summary>
        /// Gets or sets a value that represents:
        /// - The day of the month when Year is non zero,
        /// - The index of the week in the month if Year is equal to zero.
        /// </summary>
        internal int DayOrder { get; set; }

        /// <summary>
        /// Gets or sets the month.
        /// </summary>
        internal int Month { get; set; }

        /// <summary>
        /// Gets or sets the day of the week.
        /// </summary>
        internal DayOfTheWeek DayOfTheWeek { get; set; }

        /// <summary>
        /// Gets or sets the year. If Year is 0, the time change occurs every year according to a recurring pattern;
        /// otherwise, the time change occurs at the date specified by Day, Month, Year.
        /// </summary>
        internal int Year { get; set; }
    }
}