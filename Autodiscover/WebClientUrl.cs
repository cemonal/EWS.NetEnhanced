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
    using System.Xml;
    using EWS.NetEnhanced.Data;

    /// <summary>
    /// Represents the URL of the Exchange web client.
    /// </summary>
    public sealed class WebClientUrl
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="WebClientUrl"/> class.
        /// </summary>
        private WebClientUrl()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WebClientUrl"/> class.
        /// </summary>
        /// <param name="authenticationMethods">The authentication methods.</param>
        /// <param name="url">The URL.</param>
        internal WebClientUrl(string authenticationMethods, string url)
        {
            this.AuthenticationMethods = authenticationMethods;
            this.Url = url;
        }

        /// <summary>
        /// Loads WebClientUrl instance from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>WebClientUrl.</returns>
        internal static WebClientUrl LoadFromXml(EwsXmlReader reader)
        {
            WebClientUrl webClientUrl = new WebClientUrl();

            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case XmlElementNames.AuthenticationMethods:
                            webClientUrl.AuthenticationMethods = reader.ReadElementValue<string>();
                            break;
                        case XmlElementNames.Url:
                            webClientUrl.Url = reader.ReadElementValue<string>();
                            break;
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.WebClientUrl));

            return webClientUrl;
        }

        /// <summary>
        /// Gets the authentication methods.
        /// </summary>
        public string AuthenticationMethods { get; internal set; }

        /// <summary>
        /// Gets the URL.
        /// </summary>
        public string Url { get; internal set; }
    }
}