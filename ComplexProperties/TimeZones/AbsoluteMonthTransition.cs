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
    /// Represents the base class for all recurring time zone period transitions.
    /// </summary>
    internal abstract class AbsoluteMonthTransition : TimeZoneTransition
    {

        /// <summary>
        /// Initializes this transition based on the specified transition time.
        /// </summary>
        /// <param name="transitionTime">The transition time to initialize from.</param>
        internal override void InitializeFromTransitionTime(TimeZoneInfo.TransitionTime transitionTime)
        {
            base.InitializeFromTransitionTime(transitionTime);

            this.TimeOffset = transitionTime.TimeOfDay.TimeOfDay;
            this.Month = transitionTime.Month;
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            if (base.TryReadElementFromXml(reader))
            {
                return true;
            }
            else
            {
                switch (reader.LocalName)
                {
                    case XmlElementNames.TimeOffset:
                        this.TimeOffset = EwsUtilities.XSDurationToTimeSpan(reader.ReadElementValue());
                        return true;
                    case XmlElementNames.Month:
                        this.Month = reader.ReadElementValue<int>();

                        EwsUtilities.Assert(
                            this.Month > 0 && this.Month <= 12,
                            "AbsoluteMonthTransition.TryReadElementFromXml",
                            "month is not in the valid 1 - 12 range.");

                        return true;
                    default:
                        return false;
                }
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            base.WriteElementsToXml(writer);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.TimeOffset,
                EwsUtilities.TimeSpanToXSDuration(this.TimeOffset));

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.Month,
                this.Month);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AbsoluteMonthTransition"/> class.
        /// </summary>
        /// <param name="timeZoneDefinition">The time zone definition this transition belongs to.</param>
        protected AbsoluteMonthTransition(TimeZoneDefinition timeZoneDefinition) : base(timeZoneDefinition)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AbsoluteMonthTransition"/> class.
        /// </summary>
        /// <param name="timeZoneDefinition">The time zone definition this transition belongs to.</param>
        /// <param name="targetPeriod">The period the transition will target.</param>
        protected AbsoluteMonthTransition(TimeZoneDefinition timeZoneDefinition, TimeZonePeriod targetPeriod) : base(timeZoneDefinition, targetPeriod)
        {
        }

        /// <summary>
        /// Gets the time offset from midnight when the transition occurs.
        /// </summary>
        internal TimeSpan TimeOffset { get; private set; }

        /// <summary>
        /// Gets the month when the transition occurs.
        /// </summary>
        internal int Month { get; private set; }
    }
}