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
    /// Represents an error that occurred as a result of executing a rule operation. 
    /// </summary>
    public sealed class RuleError : ComplexProperty
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="RuleError"/> class.
        /// </summary>
        internal RuleError()
            : base()
        {
        }

        /// <summary>
        /// Gets the property which failed validation.
        /// </summary>
        public RuleProperty RuleProperty { get; private set; }

        /// <summary>
        /// Gets the validation error code.
        /// </summary>
        public RuleErrorCode ErrorCode { get; private set; }

        /// <summary>
        /// Gets the error message.
        /// </summary>
        public string ErrorMessage { get; private set; }

        /// <summary>
        /// Gets the value that failed validation.
        /// </summary>
        public string Value { get; private set; }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.FieldURI:
                    this.RuleProperty = reader.ReadElementValue<RuleProperty>();
                    return true;
                case XmlElementNames.ErrorCode:
                    this.ErrorCode = reader.ReadElementValue<RuleErrorCode>();
                    return true;
                case XmlElementNames.ErrorMessage:
                    this.ErrorMessage = reader.ReadElementValue();
                    return true;
                case XmlElementNames.FieldValue:
                    this.Value = reader.ReadElementValue();
                    return true;
                default:
                    return false;
            }
        }
    }
}