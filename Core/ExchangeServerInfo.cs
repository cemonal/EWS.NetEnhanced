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
    /// Represents Exchange server information.
    /// </summary>
    public sealed class ExchangeServerInfo
    {

        /// <summary>
        /// Default constructor
        /// </summary>
        internal ExchangeServerInfo()
        {
        }

        /// <summary>
        /// Parse current element to extract server information
        /// </summary>
        /// <param name="reader">EwsServiceXmlReader</param>
        /// <returns>ExchangeServerInfo</returns>
        internal static ExchangeServerInfo Parse(EwsServiceXmlReader reader)
        {
            EwsUtilities.Assert(
                                reader.HasAttributes, 
                                "ExchangeServerVersion.Parse",
                                "Current element doesn't have attributes");

            ExchangeServerInfo info = new ExchangeServerInfo
            {
                MajorVersion = reader.ReadAttributeValue<int>("MajorVersion"),
                MinorVersion = reader.ReadAttributeValue<int>("MinorVersion"),
                MajorBuildNumber = reader.ReadAttributeValue<int>("MajorBuildNumber"),
                MinorBuildNumber = reader.ReadAttributeValue<int>("MinorBuildNumber"),
                VersionString = reader.ReadAttributeValue("Version")
            };
            return info;
        }

        /// <summary>
        /// Gets the Major Exchange server version number
        /// </summary>
        public int MajorVersion { get; internal set; }

        /// <summary>
        /// Gets the Minor Exchange server version number
        /// </summary>
        public int MinorVersion { get; internal set; }

        /// <summary>
        /// Gets the Major Exchange server build number
        /// </summary>
        public int MajorBuildNumber { get; internal set; }

        /// <summary>
        /// Gets the Minor Exchange server build number
        /// </summary>
        public int MinorBuildNumber { get; internal set; }

        /// <summary>
        /// Gets the Exchange server version string (e.g. "Exchange2010")
        /// </summary>
        /// <remarks>
        /// The version is a string rather than an enum since its possible for the client to
        /// be connected to a later server for which there would be no appropriate enum value.
        /// </remarks>
        public string VersionString { get; internal set; }

        /// <summary>
        /// Override ToString method
        /// </summary>
        /// <returns>Canonical ExchangeService version string</returns>
        public override string ToString()
        {
            return string.Format(
                                 "{0:d}.{1:d2}.{2:d4}.{3:d3}",
                                 this.MajorVersion,
                                 this.MinorVersion,
                                 this.MajorBuildNumber,
                                 this.MinorBuildNumber);
        }
    }
}