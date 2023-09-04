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
    using System.Xml;

    /// <summary>
    /// Represents the response to a folder search operation.
    /// </summary>
    public sealed class FindFolderResponse : ServiceResponse
    {
        private readonly PropertySet propertySet;

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.RootFolder);

            this.Results.TotalCount = reader.ReadAttributeValue<int>(XmlAttributeNames.TotalItemsInView);
            this.Results.MoreAvailable = !reader.ReadAttributeValue<bool>(XmlAttributeNames.IncludesLastItemInRange);

            // Ignore IndexedPagingOffset attribute if MoreAvailable is false.
            this.Results.NextPageOffset = Results.MoreAvailable ? reader.ReadNullableAttributeValue<int>(XmlAttributeNames.IndexedPagingOffset) : null;

            reader.ReadStartElement(XmlNamespace.Types, XmlElementNames.Folders);
            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        Folder folder = EwsUtilities.CreateEwsObjectFromXmlElementName<Folder>(reader.Service, reader.LocalName);

                        if (folder == null)
                        {
                            reader.SkipCurrentElement();
                        }
                        else
                        {
                            folder.LoadFromXml(
                                        reader,
                                        true, /* clearPropertyBag */
                                        this.propertySet,
                                        true  /* summaryPropertiesOnly */);

                            this.Results.Folders.Add(folder);
                        }
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.Folders));
            }

            reader.ReadEndElement(XmlNamespace.Messages, XmlElementNames.RootFolder);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FindFolderResponse"/> class.
        /// </summary>
        /// <param name="propertySet">The property set from, the request.</param>
        internal FindFolderResponse(PropertySet propertySet) : base()
        {
            this.propertySet = propertySet;

            EwsUtilities.Assert(
                this.propertySet != null,
                "FindFolderResponse.ctor",
                "PropertySet should not be null");
        }

        /// <summary>
        /// Gets the results of the search operation.
        /// </summary>
        public FindFoldersResults Results { get; } = new FindFoldersResults();
    }
}