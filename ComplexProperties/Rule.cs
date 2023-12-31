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
    /// Represents a rule that automatically handles incoming messages.
    /// A rule consists of a set of conditions and exceptions that determine whether or 
    /// not a set of actions should be executed on incoming messages.
    /// </summary>
    public sealed class Rule : ComplexProperty
    {
        /// <summary>
        /// The rule ID.
        /// </summary>
        private string ruleId;

        /// <summary>
        /// The rule display name.
        /// </summary>
        private string displayName;

        /// <summary>
        /// The rule priority.
        /// </summary>
        private int priority;

        /// <summary>
        /// The rule status of enabled or not.
        /// </summary>
        private bool isEnabled;

        /// <summary>
        /// The rule status of in error or not.
        /// </summary>
        private bool isInError;

        /// <summary>
        /// Initializes a new instance of the <see cref="Rule"/> class.
        /// </summary>
        public Rule()
            : base()
        {
            //// New rule has priority as 0 by default
            this.priority = 1;
            //// New rule is enabled by default
            this.isEnabled = true;
            this.Conditions = new RulePredicates();
            this.Actions = new RuleActions();
            this.Exceptions = new RulePredicates();
        }

        /// <summary>
        /// Gets or sets the Id of this rule.
        /// </summary>
        public string Id
        {
            get
            {
                return this.ruleId;
            }

            set
            {
                this.SetFieldValue(ref this.ruleId, value);
            }
        }

        /// <summary>
        /// Gets or sets the name of this rule as it should be displayed to the user.
        /// </summary>
        public string DisplayName
        {
            get
            {
                return this.displayName;
            }

            set
            {
                this.SetFieldValue(ref this.displayName, value);
            }
        }

        /// <summary>
        /// Gets or sets the priority of this rule, which determines its execution order.
        /// </summary>
        public int Priority
        {
            get
            {
                return this.priority;
            }

            set
            {
                this.SetFieldValue(ref this.priority, value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this rule is enabled.
        /// </summary>
        public bool IsEnabled
        {
            get
            {
                return this.isEnabled;
            }

            set
            {
                this.SetFieldValue(ref this.isEnabled, value);
            }
        }

        /// <summary>
        /// Gets a value indicating whether this rule can be modified via EWS. 
        /// If IsNotSupported is true, the rule cannot be modified via EWS.
        /// </summary>
        public bool IsNotSupported { get; private set; }

        /// <summary>
        /// Gets or sets a value indicating whether this rule has errors. A rule that is in error 
        /// cannot be processed unless it is updated and the error is corrected.
        /// </summary>
        public bool IsInError
        {
            get
            {
                return this.isInError;
            }

            set
            {
                this.SetFieldValue(ref this.isInError, value);
            }
        }

        /// <summary>
        /// Gets the conditions that determine whether or not this rule should be
        /// executed against incoming messages.
        /// </summary>
        public RulePredicates Conditions { get; }

        /// <summary>
        /// Gets the actions that should be executed against incoming messages if the
        /// conditions evaluate as true.
        /// </summary>
        public RuleActions Actions { get; }

        /// <summary>
        /// Gets the exceptions that determine if this rule should be skipped even if 
        /// its conditions evaluate to true.
        /// </summary>
        public RulePredicates Exceptions { get; }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.DisplayName:
                    this.displayName = reader.ReadElementValue();
                    return true;
                case XmlElementNames.RuleId:
                    this.ruleId = reader.ReadElementValue();
                    return true;
                case XmlElementNames.Priority:
                    this.priority = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.IsEnabled:
                    this.isEnabled = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsNotSupported:
                    this.IsNotSupported = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsInError:
                    this.isInError = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.Conditions:
                    this.Conditions.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.Actions:
                    this.Actions.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.Exceptions:
                    this.Exceptions.LoadFromXml(reader, reader.LocalName);
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (!string.IsNullOrEmpty(this.Id))
            {
                writer.WriteElementValue(
                    XmlNamespace.Types, 
                    XmlElementNames.RuleId, 
                    this.Id);
            }

            writer.WriteElementValue(
                XmlNamespace.Types, 
                XmlElementNames.DisplayName, 
                this.DisplayName);
            writer.WriteElementValue(
                XmlNamespace.Types, 
                XmlElementNames.Priority, 
                this.Priority);
            writer.WriteElementValue(
                XmlNamespace.Types, 
                XmlElementNames.IsEnabled, 
                this.IsEnabled);
            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.IsInError,
                this.IsInError);
            this.Conditions.WriteToXml(writer, XmlElementNames.Conditions);
            this.Exceptions.WriteToXml(writer, XmlElementNames.Exceptions);
            this.Actions.WriteToXml(writer, XmlElementNames.Actions);
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void InternalValidate()
        {
            base.InternalValidate();
            EwsUtilities.ValidateParam(this.displayName, "DisplayName");
            EwsUtilities.ValidateParam(this.Conditions, "Conditions");
            EwsUtilities.ValidateParam(this.Exceptions, "Exceptions");
            EwsUtilities.ValidateParam(this.Actions, "Actions");
        }
    }
}