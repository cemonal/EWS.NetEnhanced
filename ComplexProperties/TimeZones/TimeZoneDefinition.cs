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
    using System.Collections.Generic;

    /// <summary>
    /// Represents a time zone as defined by the EWS schema.
    /// </summary>
    internal class TimeZoneDefinition : ComplexProperty
    {
        /// <summary>
        /// Prefix for generated ids.
        /// </summary>
        private const string NoIdPrefix = "NoId_";
        private readonly List<TimeZoneTransition> transitions = new List<TimeZoneTransition>();

        /// <summary>
        /// Compares the transitions.
        /// </summary>
        /// <param name="x">The first transition.</param>
        /// <param name="y">The second transition.</param>
        /// <returns>A negative number if x is less than y, 0 if x and y are equal, a positive number if x is greater than y.</returns>
        private int CompareTransitions(TimeZoneTransition x, TimeZoneTransition y)
        {
            if (x == y)
            {
                return 0;
            }
            else if (x.GetType() == typeof(TimeZoneTransition))
            {
                return -1;
            }
            else if (y.GetType() == typeof(TimeZoneTransition))
            {
                return 1;
            }
            else
            {
                AbsoluteDateTransition firstTransition = (AbsoluteDateTransition)x;
                AbsoluteDateTransition secondTransition = (AbsoluteDateTransition)y;

                return DateTime.Compare(firstTransition.DateTime, secondTransition.DateTime);
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TimeZoneDefinition"/> class.
        /// </summary>
        internal TimeZoneDefinition()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TimeZoneDefinition"/> class.
        /// </summary>
        /// <param name="timeZoneInfo">The time zone info used to initialize this definition.</param>
        internal TimeZoneDefinition(TimeZoneInfo timeZoneInfo)
            : this()
        {
            this.Id = timeZoneInfo.Id;
            this.Name = timeZoneInfo.DisplayName;

            // TimeZoneInfo only supports one standard period, which bias is the time zone's base
            // offset to UTC.
            TimeZonePeriod standardPeriod = new TimeZonePeriod
            {
                Id = TimeZonePeriod.StandardPeriodId,
                Name = TimeZonePeriod.StandardPeriodName,
                Bias = -timeZoneInfo.BaseUtcOffset
            };

            TimeZoneInfo.AdjustmentRule[] adjustmentRules = timeZoneInfo.GetAdjustmentRules();

            TimeZoneTransition transitionToStandardPeriod = new TimeZoneTransition(this, standardPeriod);

            if (adjustmentRules.Length == 0)
            {
                this.Periods.Add(standardPeriod.Id, standardPeriod);

                // If the time zone info doesn't support Daylight Saving Time, we just need to
                // create one transition to one group with one transition to the standard period.
                TimeZoneTransitionGroup transitionGroup = new TimeZoneTransitionGroup(this, "0");
                transitionGroup.Transitions.Add(transitionToStandardPeriod);

                this.TransitionGroups.Add(transitionGroup.Id, transitionGroup);

                TimeZoneTransition initialTransition = new TimeZoneTransition(this, transitionGroup);

                this.transitions.Add(initialTransition);
            }
            else
            {
                for (int i = 0; i < adjustmentRules.Length; i++)
                {
                    TimeZoneTransitionGroup transitionGroup = new TimeZoneTransitionGroup(this, this.TransitionGroups.Count.ToString());
                    transitionGroup.InitializeFromAdjustmentRule(adjustmentRules[i], standardPeriod);

                    this.TransitionGroups.Add(transitionGroup.Id, transitionGroup);

                    TimeZoneTransition transition;

                    if (i == 0)
                    {
                        // If the first adjustment rule's start date in not undefined (DateTime.MinValue)
                        // we need to add a dummy group with a single, simple transition to the Standard
                        // period and a group containing the transitions mapping to the adjustment rule.
                        if (adjustmentRules[i].DateStart > DateTime.MinValue.Date)
                        {
                            TimeZoneTransition transitionToDummyGroup = new TimeZoneTransition(
                                this,
                                this.CreateTransitionGroupToPeriod(standardPeriod));

                            this.transitions.Add(transitionToDummyGroup);

                            AbsoluteDateTransition absoluteDateTransition = new AbsoluteDateTransition(this, transitionGroup)
                            {
                                DateTime = adjustmentRules[i].DateStart
                            };

                            transition = absoluteDateTransition;
                            this.Periods.Add(standardPeriod.Id, standardPeriod);
                        }
                        else
                        {
                            transition = new TimeZoneTransition(this, transitionGroup);
                        }
                    }
                    else
                    {
                        AbsoluteDateTransition absoluteDateTransition = new AbsoluteDateTransition(this, transitionGroup)
                        {
                            DateTime = adjustmentRules[i].DateStart
                        };

                        transition = absoluteDateTransition;
                    }

                    this.transitions.Add(transition);
                }

                // If the last adjustment rule's end date is not undefined (DateTime.MaxValue),
                // we need to create another absolute date transition that occurs the date after
                // the last rule's end date. We target this additional transition to a group that
                // contains a single simple transition to the Standard period.
                DateTime lastAdjustmentRuleEndDate = adjustmentRules[adjustmentRules.Length - 1].DateEnd;

                if (lastAdjustmentRuleEndDate < DateTime.MaxValue.Date)
                {
                    AbsoluteDateTransition transitionToDummyGroup = new AbsoluteDateTransition(
                        this,
                        this.CreateTransitionGroupToPeriod(standardPeriod))
                    {
                        DateTime = lastAdjustmentRuleEndDate.AddDays(1)
                    };

                    this.transitions.Add(transitionToDummyGroup);
                }
            }
        }

        /// <summary>
        /// Adds a transition group with a single transition to the specified period.
        /// </summary>
        /// <param name="timeZonePeriod">The time zone period.</param>
        /// <returns>A TimeZoneTransitionGroup.</returns>
        private TimeZoneTransitionGroup CreateTransitionGroupToPeriod(TimeZonePeriod timeZonePeriod)
        {
            TimeZoneTransition transitionToPeriod = new TimeZoneTransition(this, timeZonePeriod);

            TimeZoneTransitionGroup transitionGroup = new TimeZoneTransitionGroup(this, this.TransitionGroups.Count.ToString());
            transitionGroup.Transitions.Add(transitionToPeriod);

            this.TransitionGroups.Add(transitionGroup.Id, transitionGroup);

            return transitionGroup;
        }

        /// <summary>
        /// Reads the attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
        {
            this.Name = reader.ReadAttributeValue(XmlAttributeNames.Name);
            this.Id = reader.ReadAttributeValue(XmlAttributeNames.Id);

            // EWS can return a TimeZone definition with no Id. Generate a new Id in this case.
            if (string.IsNullOrEmpty(this.Id))
            {
                string nameValue = string.IsNullOrEmpty(this.Name) ? string.Empty : this.Name;
                this.Id = NoIdPrefix + Math.Abs(nameValue.GetHashCode()).ToString();
            }
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            // The Name attribute is only supported in Exchange 2010 and above.
            if (writer.Service.RequestedServerVersion != ExchangeVersion.Exchange2007_SP1)
            {
                writer.WriteAttributeValue(XmlAttributeNames.Name, this.Name);
            }

            writer.WriteAttributeValue(XmlAttributeNames.Id, this.Id);
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
                case XmlElementNames.Periods:
                    do
                    {
                        reader.Read();

                        if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Period))
                        {
                            TimeZonePeriod period = new TimeZonePeriod();
                            period.LoadFromXml(reader);

                            // OM:1648848 Bad timezone data from clients can include duplicate rules
                            // for one year, with duplicate ID. In that case, let the first one win.
                            if (!this.Periods.ContainsKey(period.Id))
                            {
                                this.Periods.Add(period.Id, period);
                            }
                            else
                            {
                                reader.Service.TraceMessage(
                                    TraceFlags.EwsTimeZones,
                                    string.Format(
                                        "An entry with the same key (Id) '{0}' already exists in Periods. Cannot add another one. Existing entry: [Name='{1}', Bias='{2}']. Entry to skip: [Name='{3}', Bias='{4}'].",
                                        period.Id,
                                        this.Periods[period.Id].Name,
                                        this.Periods[period.Id].Bias,
                                        period.Name,
                                        period.Bias));
                            }
                        }
                    }
                    while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.Periods));

                    return true;
                case XmlElementNames.TransitionsGroups:
                    do
                    {
                        reader.Read();

                        if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.TransitionsGroup))
                        {
                            TimeZoneTransitionGroup transitionGroup = new TimeZoneTransitionGroup(this);

                            transitionGroup.LoadFromXml(reader);

                            this.TransitionGroups.Add(transitionGroup.Id, transitionGroup);
                        }
                    }
                    while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.TransitionsGroups));

                    return true;
                case XmlElementNames.Transitions:
                    do
                    {
                        reader.Read();

                        if (reader.IsStartElement())
                        {
                            TimeZoneTransition transition = TimeZoneTransition.Create(this, reader.LocalName);

                            transition.LoadFromXml(reader);

                            this.transitions.Add(transition);
                        }
                    }
                    while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.Transitions));

                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsServiceXmlReader reader)
        {
            this.LoadFromXml(reader, XmlElementNames.TimeZoneDefinition);

            this.transitions.Sort(this.CompareTransitions);
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            // We only emit the full time zone definition against Exchange 2010 servers and above.
            if (writer.Service.RequestedServerVersion != ExchangeVersion.Exchange2007_SP1)
            {
                if (this.Periods.Count > 0)
                {
                    writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Periods);

                    foreach (KeyValuePair<string, TimeZonePeriod> keyValuePair in this.Periods)
                    {
                        keyValuePair.Value.WriteToXml(writer);
                    }

                    writer.WriteEndElement(); // Periods
                }

                if (this.TransitionGroups.Count > 0)
                {
                    writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.TransitionsGroups);

                    foreach (KeyValuePair<string, TimeZoneTransitionGroup> keyValuePair in this.TransitionGroups)
                    {
                        keyValuePair.Value.WriteToXml(writer);
                    }

                    writer.WriteEndElement(); // TransitionGroups
                }

                if (this.transitions.Count > 0)
                {
                    writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Transitions);

                    foreach (TimeZoneTransition transition in this.transitions)
                    {
                        transition.WriteToXml(writer);
                    }

                    writer.WriteEndElement(); // Transitions
                }
            }
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer)
        {
            this.WriteToXml(writer, XmlElementNames.TimeZoneDefinition);
        }

        /// <summary>
        /// Validates this time zone definition.
        /// </summary>
        internal void Validate()
        {
            // The definition must have at least one period, one transition group and one transition,
            // and there must be as many transitions as there are transition groups.
            if (this.Periods.Count < 1 || this.transitions.Count < 1 || this.TransitionGroups.Count < 1 ||
                this.TransitionGroups.Count != this.transitions.Count)
            {
                throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
            }

            // The first transition must be of type TimeZoneTransition.
            if (this.transitions[0].GetType() != typeof(TimeZoneTransition))
            {
                throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
            }

            // All transitions must be to transition groups and be either TimeZoneTransition or
            // AbsoluteDateTransition instances.
            foreach (TimeZoneTransition transition in this.transitions)
            {
                Type transitionType = transition.GetType();

                if (transitionType != typeof(TimeZoneTransition) && transitionType != typeof(AbsoluteDateTransition))
                {
                    throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
                }

                if (transition.TargetGroup == null)
                {
                    throw new ServiceLocalException(Strings.InvalidOrUnsupportedTimeZoneDefinition);
                }
            }

            // All transition groups must be valid.
            foreach (TimeZoneTransitionGroup transitionGroup in this.TransitionGroups.Values)
            {
                transitionGroup.Validate();
            }
        }

        /// <summary>
        /// Converts this time zone definition into a TimeZoneInfo structure.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>A TimeZoneInfo representing the same time zone as this definition.</returns>
        internal TimeZoneInfo ToTimeZoneInfo(ExchangeService service)
        {
            this.Validate();

            TimeZoneInfo result;

            // Retrieve the base offset to UTC, standard and daylight display names from
            // the last transition group, which is the one that currently applies given that
            // transitions are ordered chronologically.
            TimeZoneTransitionGroup.CustomTimeZoneCreateParams creationParams =
                this.transitions[this.transitions.Count - 1].TargetGroup.GetCustomTimeZoneCreationParams();

            List<TimeZoneInfo.AdjustmentRule> adjustmentRules = new List<TimeZoneInfo.AdjustmentRule>();

            DateTime startDate = DateTime.MinValue;
            DateTime endDate;
            DateTime effectiveEndDate;

            for (int i = 0; i < this.transitions.Count; i++)
            {
                if (i < this.transitions.Count - 1)
                {
                    endDate = (this.transitions[i + 1] as AbsoluteDateTransition).DateTime;
                    effectiveEndDate = endDate.AddDays(-1);
                }
                else
                {
                    endDate = DateTime.MaxValue;
                    effectiveEndDate = endDate;
                }

                // OM:1648848 Due to bad timezone data from clients the 
                // startDate may not always come before the effectiveEndDate
                if (startDate < effectiveEndDate)
                {
                    TimeZoneInfo.AdjustmentRule adjustmentRule = this.transitions[i].TargetGroup.CreateAdjustmentRule(startDate, effectiveEndDate);

                    if (adjustmentRule != null)
                    {
                        adjustmentRules.Add(adjustmentRule);
                    }

                    startDate = endDate;
                }
                else
                {
                    service.TraceMessage(
                        TraceFlags.EwsTimeZones,
                            string.Format(
                                "The startDate '{0}' is not before the effectiveEndDate '{1}'. Will skip creating adjustment rule.",
                                startDate,
                                effectiveEndDate));
                }
            }

            if (adjustmentRules.Count == 0)
            {
                // If there are no adjustment rule, the time zone does not support Daylight
                // saving time.
                result = TimeZoneInfo.CreateCustomTimeZone(
                    this.Id,
                    creationParams.BaseOffsetToUtc,
                    this.Name,
                    creationParams.StandardDisplayName);
            }
            else
            {
                result = TimeZoneInfo.CreateCustomTimeZone(
                    this.Id,
                    creationParams.BaseOffsetToUtc,
                    this.Name,
                    creationParams.StandardDisplayName,
                    creationParams.DaylightDisplayName,
                    adjustmentRules.ToArray());
            }

            return result;
        }

        /// <summary>
        /// Gets or sets the name of this time zone definition.
        /// </summary>
        internal string Name { get; set; }

        /// <summary>
        /// Gets or sets the Id of this time zone definition.
        /// </summary>
        internal string Id { get; set; }

        /// <summary>
        /// Gets the periods associated with this time zone definition, indexed by Id.
        /// </summary>
        internal Dictionary<string, TimeZonePeriod> Periods { get; } = new Dictionary<string, TimeZonePeriod>();

        /// <summary>
        /// Gets the transition groups associated with this time zone definition, indexed by Id.
        /// </summary>
        internal Dictionary<string, TimeZoneTransitionGroup> TransitionGroups { get; } = new Dictionary<string, TimeZoneTransitionGroup>();
    }
}