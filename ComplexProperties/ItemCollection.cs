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
    using System.ComponentModel;
    using System.Xml;

    /// <summary>
    /// Represents a collection of items.
    /// </summary>
    /// <typeparam name="TItem">The type of item the collection contains.</typeparam>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class ItemCollection<TItem> : ComplexProperty, IEnumerable<TItem> where TItem : Item
    {
        private readonly List<TItem> items = new List<TItem>();

        /// <summary>
        /// Initializes a new instance of the <see cref="ItemCollection&lt;TItem&gt;"/> class.
        /// </summary>
        internal ItemCollection() : base() { }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="xmlElementName">Name of the local element.</param>
        internal override void LoadFromXml(EwsServiceXmlReader reader, string xmlElementName)
        {
            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, xmlElementName);
            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        if (!(EwsUtilities.CreateEwsObjectFromXmlElementName<Item>(
                            reader.Service,
                            reader.LocalName) is TItem item))
                        {
                            reader.SkipCurrentElement();
                        }
                        else
                        {
                            item.LoadFromXml(reader, true /* clearPropertyBag */);

                            this.items.Add(item);
                        }
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Types, xmlElementName));
            }
        }

        /// <summary>
        /// Gets the total number of items in the collection.
        /// </summary>
        public int Count => this.items.Count;

        /// <summary>
        /// Gets the item at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index of the item to get.</param>
        /// <returns>The item at the specified index.</returns>
        public TItem this[int index]
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

        #region IEnumerable<TItem> Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<TItem> GetEnumerator()
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