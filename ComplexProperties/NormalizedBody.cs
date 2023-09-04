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
    /// Represents the normalized body of an item - the HTML fragment representation of the body.
    /// </summary>
    public sealed class NormalizedBody : ComplexProperty
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="NormalizedBody"/> class.
        /// </summary>
        internal NormalizedBody()
        {
        }

        /// <summary>
        /// Defines an implicit conversion of NormalizedBody into a string.
        /// </summary>
        /// <param name="messageBody">The NormalizedBody to convert to a string.</param>
        /// <returns>A string containing the text of the UniqueBody.</returns>
        public static implicit operator string(NormalizedBody messageBody)
        {
            EwsUtilities.ValidateParam(messageBody, "messageBody");
            return messageBody.Text;
        }

        /// <summary>
        /// Reads attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
        {
            this.BodyType = reader.ReadAttributeValue<BodyType>(XmlAttributeNames.BodyType);

            string attributeValue = reader.ReadAttributeValue(XmlAttributeNames.IsTruncated);
            if (!string.IsNullOrEmpty(attributeValue))
            {
                this.IsTruncated = bool.Parse(attributeValue);
            }
        }

        /// <summary>
        /// Reads text value from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadTextValueFromXml(EwsServiceXmlReader reader)
        {
            this.Text = reader.ReadValue();
        }

        /// <summary>
        /// Writes attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.BodyType, this.BodyType);
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (!string.IsNullOrEmpty(this.Text))
            {
                writer.WriteValue(this.Text, XmlElementNames.NormalizedBody);
            }
        }

        /// <summary>
        /// Gets the type of the normalized body's text.
        /// </summary>
        public BodyType BodyType { get; internal set; }

        /// <summary>
        /// Gets the text of the normalized body.
        /// </summary>
        public string Text { get; internal set; }

        /// <summary>
        /// Gets whether the body is truncated.
        /// </summary>
        public bool IsTruncated { get; internal set; }

        #region Object method overrides
        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override string ToString()
        {
            return (this.Text == null) ? string.Empty : this.Text;
        }
        #endregion
    }
}