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
    /// Represents an Internet message header.
    /// </summary>
    public sealed class InternetMessageHeader : ComplexProperty
    {
        private string name;
        private string value;

        /// <summary>
        /// Initializes a new instance of the <see cref="InternetMessageHeader"/> class.
        /// </summary>
        internal InternetMessageHeader()
        {
        }

        /// <summary>
        /// Reads the attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
        {
            this.name = reader.ReadAttributeValue(XmlAttributeNames.HeaderName);
        }

        /// <summary>
        /// Reads the text value from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadTextValueFromXml(EwsServiceXmlReader reader)
        {
            this.value = reader.ReadValue();
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.HeaderName, this.Name);
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteValue(this.Value, this.Name);
        }

        /// <summary>
        /// Obtains a string representation of the header.
        /// </summary>
        /// <returns>The string representation of the header.</returns>
        public override string ToString()
        {
            return string.Format("{0}={1}", this.Name, this.Value);
        }

        /// <summary>
        /// The name of the header.
        /// </summary>
        public string Name
        {
            get { return this.name; }
            set { this.SetFieldValue(ref this.name, value); }
        }

        /// <summary>
        /// The value of the header.
        /// </summary>
        public string Value
        {
            get { return this.value; }
            set { this.SetFieldValue(ref this.value, value); }
        }
    }
}