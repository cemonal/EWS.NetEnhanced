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
    /// Represents the view settings in a folder search operation.
    /// </summary>
    public sealed class SeekToConditionItemView : ViewBase
    {
        private int pageSize;
        private SearchFilter condition;
        private ServiceObjectType serviceObjType;

        /// <summary>
        /// Gets the type of service object this view applies to.
        /// </summary>
        /// <returns>A ServiceObjectType value.</returns>
        internal override ServiceObjectType GetServiceObjectType()
        {
            return this.serviceObjType;
        }

        /// <summary>
        /// Sets the type of service object this view applies to.
        /// </summary>
        /// <param name="objType">Service object type</param>
        internal void SetServiceObjectType(ServiceObjectType objType)
        {
            this.serviceObjType = objType;
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            if (this.serviceObjType == ServiceObjectType.Item)
            {
                writer.WriteAttributeValue(XmlAttributeNames.Traversal, this.Traversal);
            }
        }

        /// <summary>
        /// Gets the name of the view XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetViewXmlElementName()
        {
            return XmlElementNames.SeekToConditionPageItemView;
        }

        /// <summary>
        /// Write to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void InternalWriteViewToXml(EwsServiceXmlWriter writer)
        {
            base.InternalWriteViewToXml(writer);

            writer.WriteAttributeValue(XmlAttributeNames.BasePoint, this.OffsetBasePoint);

            if (this.Condition != null)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Condition);
                this.Condition.WriteToXml(writer);
                writer.WriteEndElement(); // Restriction
            }
        }

        /// <summary>
        /// Internals the write search settings to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="groupBy">The group by.</param>
        internal override void InternalWriteSearchSettingsToXml(EwsServiceXmlWriter writer, Grouping groupBy)
        {
            if (groupBy != null)
            {
                groupBy.WriteToXml(writer);
            }
        }

        /// <summary>
        /// Gets the maximum number of items or folders the search operation should return.
        /// </summary>
        /// <returns>The maximum number of items that should be returned by the search operation.</returns>
        internal override int? GetMaxEntriesReturned()
        {
            return this.PageSize;
        }

        /// <summary>
        /// Writes OrderBy property to XML.
        /// </summary>
        /// <param name="writer">The writer</param>
        internal override void WriteOrderByToXml(EwsServiceXmlWriter writer)
        {
            this.OrderBy.WriteToXml(writer, XmlElementNames.SortOrder);
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="groupBy">The group by clause.</param>
        internal override void WriteToXml(EwsServiceXmlWriter writer, Grouping groupBy)
        {
            if (this.serviceObjType == ServiceObjectType.Item)
            {
                this.GetPropertySetOrDefault().WriteToXml(writer, this.GetServiceObjectType());
            }

            writer.WriteStartElement(XmlNamespace.Messages, this.GetViewXmlElementName());

            this.InternalWriteViewToXml(writer);

            writer.WriteEndElement(); // this.GetViewXmlElementName()

            this.InternalWriteSearchSettingsToXml(writer, groupBy);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SeekToConditionItemView"/> class.
        /// </summary>
        /// <param name="condition">Condition to be used when seeking.</param>
        /// <param name="pageSize">The maximum number of elements the search operation should return.</param>
        public SeekToConditionItemView(SearchFilter condition, int pageSize)
        {
            this.Condition = condition;
            this.PageSize = pageSize;
            this.serviceObjType = ServiceObjectType.Item;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SeekToConditionItemView"/> class.
        /// </summary>
        /// <param name="condition">Condition to be used when seeking.</param>
        /// <param name="pageSize">The maximum number of elements the search operation should return.</param>
        /// <param name="offsetBasePoint">The base point of the offset.</param>
        public SeekToConditionItemView(
            SearchFilter condition,
            int pageSize,
            OffsetBasePoint offsetBasePoint)
            : this(condition, pageSize)
        {
            this.OffsetBasePoint = offsetBasePoint;
        }

        /// <summary>
        /// The maximum number of items or folders the search operation should return.
        /// </summary>
        public int PageSize
        {
            get
            {
                return this.pageSize;
            }

            set
            {
                if (value <= 0)
                {
                    throw new ArgumentException(Strings.ValueMustBeGreaterThanZero);
                }

                this.pageSize = value;
            }
        }

        /// <summary>
        /// Gets or sets the base point of the offset.
        /// </summary>
        public OffsetBasePoint OffsetBasePoint { get; set; } = OffsetBasePoint.Beginning;

        /// <summary>
        /// Gets or sets the condition for seek. Available search filter classes include SearchFilter.IsEqualTo,
        /// SearchFilter.ContainsSubstring and SearchFilter.SearchFilterCollection. If SearchFilter
        /// is null, no search filters are applied.
        /// </summary>
        public SearchFilter Condition
        {
            get 
            { 
                return this.condition; 
            }

            set 
            {
                if (value == null)
                {
                    throw new ArgumentNullException("Condition");
                }

                this.condition = value; 
            }
        }

        /// <summary>
        /// Gets or sets the search traversal mode. Defaults to ItemTraversal.Shallow.
        /// </summary>
        public ItemTraversal Traversal { get; set; }

        /// <summary>
        /// Gets the properties against which the returned items should be ordered.
        /// </summary>
        public OrderByCollection OrderBy { get; } = new OrderByCollection();
    }
}