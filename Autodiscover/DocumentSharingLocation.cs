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

namespace EWS.NetEnhanced.Autodiscover
{
    using System.Collections.Generic;
    using System.Xml;
    using EWS.NetEnhanced.Data;

    /// <summary>
    /// Represents a sharing location.
    /// </summary>
    public sealed class DocumentSharingLocation
    {

        /// <summary>
        /// Gets the URL of the web service to use to manipulate 
        /// documents at the sharing location.
        /// </summary>
        public string ServiceUrl { get; private set; }

        /// <summary>
        /// Gets the URL of the sharing location (for viewing the 
        /// contents in a web browser).
        /// </summary>
        public string LocationUrl { get; private set; }

        /// <summary>
        /// Gets the display name of the location.
        /// </summary>
        public string DisplayName { get; private set; }

        /// <summary>
        /// Gets the space-separated list of file extensions that are 
        /// allowed at the location.
        /// </summary>
        /// <remarks>
        /// Example:  "docx pptx xlsx"
        /// </remarks>
        public IEnumerable<string> SupportedFileExtensions { get; private set; }

        /// <summary>
        /// Gets a flag indicating whether external users (outside the 
        /// enterprise/tenant) can view documents at the location.
        /// </summary>
        public bool ExternalAccessAllowed { get; private set; }

        /// <summary>
        /// Gets a flag indicating whether anonymous users can view
        /// documents at the location.
        /// </summary>
        public bool AnonymousAccessAllowed { get; private set; }

        /// <summary>
        /// Gets a flag indicating whether the user can modify 
        /// permissions for documents at the location.
        /// </summary>
        /// <remarks>
        /// This will be true for the user's "My Site," for example. However,
        /// documents at team and project sites will typically be ACLed by the
        /// site owner, so the user will not be able to modify permissions. 
        /// This will most likely by false even if the caller is the owner,
        /// to avoid surprises.  They should go to SharePoint to modify
        /// permissions for team and project sites.
        /// </remarks>
        public bool CanModifyPermissions { get; private set; }

        /// <summary>
        /// Gets a flag indicating whether this location is the user's
        /// default location.  This will generally be their My Site.
        /// </summary>
        public bool IsDefault { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentSharingLocation"/> class.
        /// </summary>
        private DocumentSharingLocation()
        {
        }

        /// <summary>
        /// Loads DocumentSharingLocation instance from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>DocumentSharingLocation.</returns>
        internal static DocumentSharingLocation LoadFromXml(EwsXmlReader reader)
        {
            DocumentSharingLocation location = new DocumentSharingLocation();

            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case XmlElementNames.ServiceUrl:
                            location.ServiceUrl = reader.ReadElementValue<string>();
                            break;

                        case XmlElementNames.LocationUrl:
                            location.LocationUrl = reader.ReadElementValue<string>();
                            break;

                        case XmlElementNames.DisplayName:
                            location.DisplayName = reader.ReadElementValue<string>();
                            break;

                        case XmlElementNames.SupportedFileExtensions:
                            List<string> fileExtensions = new List<string>();
                            reader.Read();
                            while (reader.IsStartElement(XmlNamespace.Autodiscover, XmlElementNames.FileExtension))
                            {                                
                                string extension = reader.ReadElementValue<string>();
                                fileExtensions.Add(extension);
                                reader.Read();
                            }
                            
                            location.SupportedFileExtensions = fileExtensions;
                            break;

                        case XmlElementNames.ExternalAccessAllowed:
                            location.ExternalAccessAllowed = reader.ReadElementValue<bool>();
                            break;

                        case XmlElementNames.AnonymousAccessAllowed:
                            location.AnonymousAccessAllowed = reader.ReadElementValue<bool>();
                            break;

                        case XmlElementNames.CanModifyPermissions:
                            location.CanModifyPermissions = reader.ReadElementValue<bool>();
                            break;

                        case XmlElementNames.IsDefault:
                            location.IsDefault = reader.ReadElementValue<bool>();
                            break;

                        default:
                            break;
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.DocumentSharingLocation));

            return location;
        }
    }
}