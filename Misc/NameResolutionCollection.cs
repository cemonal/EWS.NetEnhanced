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
    /// Represents a list of suggested name resolutions.
    /// </summary>
    public sealed class NameResolutionCollection : IEnumerable<NameResolution>
    {
        private readonly List<NameResolution> items = new List<NameResolution>();

        /// <summary>
        /// Initializes a new instance of the <see cref="NameResolutionCollection"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal NameResolutionCollection(ExchangeService service)
        {
            EwsUtilities.Assert(
                service != null,
                "NameResolutionSet.ctor",
                "service is null.");

            this.Session = service;
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsServiceXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.ResolutionSet);

            int totalItemsInView = reader.ReadAttributeValue<int>(XmlAttributeNames.TotalItemsInView);
            this.IncludesAllResolutions = reader.ReadAttributeValue<bool>(XmlAttributeNames.IncludesLastItemInRange);

            for (int i = 0; i < totalItemsInView; i++)
            {
                NameResolution nameResolution = new NameResolution(this);

                nameResolution.LoadFromXml(reader);

                this.items.Add(nameResolution);
            }

            reader.ReadEndElement(XmlNamespace.Messages, XmlElementNames.ResolutionSet);
        }

        /// <summary>
        /// Gets the session.
        /// </summary>
        /// <value>The session.</value>
        internal ExchangeService Session { get; }

        /// <summary>
        /// Gets the total number of elements in the list.
        /// </summary>
        public int Count => this.items.Count;

        /// <summary>
        /// Gets a value indicating whether more suggested resolutions are available. ResolveName only returns
        /// a maximum of 100 name resolutions. When IncludesAllResolutions is false, there were more than 100
        /// matching names on the server. To narrow the search, provide a more precise name to ResolveName.
        /// </summary>
        public bool IncludesAllResolutions { get; private set; }

        /// <summary>
        /// Gets the name resolution at the specified index.
        /// </summary>
        /// <param name="index">The index of the name resolution to get.</param>
        /// <returns>The name resolution at the speicfied index.</returns>
        public NameResolution this[int index]
        {
            get
            {
                if (index < 0 || index >= this.Count)
                {
                    throw new ArgumentOutOfRangeException("index", Strings.IndexIsOutOfRange);
                }

                return this.items[index];
            }
        }

        #region IEnumerable<NameResolution> Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<NameResolution> GetEnumerator()
        {
            return this.items.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.items.GetEnumerator();
        }

        #endregion
    }
}