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
    using System.Diagnostics.CodeAnalysis;
    using System.Linq;
    using System.Reflection;

    using PropertyDefinitionDictionary = LazyMember<System.Collections.Generic.Dictionary<string, PropertyDefinitionBase>>;
    using SchemaTypeList = LazyMember<System.Collections.Generic.List<System.Type>>;

    /// <summary>
    /// Represents the base class for all item and folder schemas.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class ServiceObjectSchema : IEnumerable<PropertyDefinition>
    {
        private static readonly object lockObject = new object();

        /// <summary>
        /// List of all schema types.
        /// </summary>
        /// <remarks>
        /// If you add a new ServiceObject subclass that has an associated schema, add the schema type
        /// to the list below.
        /// </remarks>
        private static readonly SchemaTypeList allSchemaTypes = new SchemaTypeList(
            delegate ()
            {
                List<Type> typeList = new List<Type>
                {
                    typeof(AppointmentSchema),
                    typeof(CalendarResponseObjectSchema),
                    typeof(CancelMeetingMessageSchema),
                    typeof(ContactGroupSchema),
                    typeof(ContactSchema),
                    typeof(ConversationSchema),
                    typeof(EmailMessageSchema),
                    typeof(FolderSchema),
                    typeof(ItemSchema),
                    typeof(MeetingMessageSchema),
                    typeof(MeetingRequestSchema),
                    typeof(MeetingCancellationSchema),
                    typeof(MeetingResponseSchema),
                    typeof(PersonaSchema),
                    typeof(PostItemSchema),
                    typeof(PostReplySchema),
                    typeof(ResponseMessageSchema),
                    typeof(ResponseObjectSchema),
                    typeof(ServiceObjectSchema),
                    typeof(SearchFolderSchema),
                    typeof(TaskSchema)
                };

#if DEBUG
                // Verify that all Schema types in the Managed API assembly have been included.
                var missingTypes = from type in Assembly.GetExecutingAssembly().GetTypes()
                                   where type.IsSubclassOf(typeof(ServiceObjectSchema)) && !typeList.Contains(type)
                                   select type;
                if (missingTypes.Count() > 0)
                {
                    throw new ServiceLocalException("SchemaTypeList does not include all defined schema types.");
                }
#endif

                return typeList;
            });

        /// <summary>
        /// Dictionary of all property definitions.
        /// </summary>
        private static readonly PropertyDefinitionDictionary allSchemaProperties = new PropertyDefinitionDictionary(
            delegate ()
            {
                Dictionary<string, PropertyDefinitionBase> propDefDictionary = new Dictionary<string, PropertyDefinitionBase>();
                foreach (Type type in allSchemaTypes.Member)
                {
                    AddSchemaPropertiesToDictionary(type, propDefDictionary);
                }
                return propDefDictionary;
            });

        /// <summary>
        /// Delegate that takes a property definition and matching static field info.
        /// </summary>
        /// <param name="propertyDefinition">Property definition.</param>
        /// <param name="fieldInfo">Field info.</param>
        internal delegate void PropertyFieldInfoDelegate(PropertyDefinition propertyDefinition, FieldInfo fieldInfo);

        /// <summary>
        /// Call delegate for each public static PropertyDefinition field in type.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <param name="propFieldDelegate">The property field delegate.</param>
        internal static void ForeachPublicStaticPropertyFieldInType(Type type, PropertyFieldInfoDelegate propFieldDelegate)
        {
            FieldInfo[] fieldInfos = type.GetFields(BindingFlags.Static | BindingFlags.Public | BindingFlags.DeclaredOnly);

            foreach (FieldInfo fieldInfo in fieldInfos)
            {
                if (fieldInfo.FieldType == typeof(PropertyDefinition) || fieldInfo.FieldType.IsSubclassOf(typeof(PropertyDefinition)))
                {
                    PropertyDefinition propertyDefinition = (PropertyDefinition)fieldInfo.GetValue(null);
                    propFieldDelegate(propertyDefinition, fieldInfo);
                }
            }
        }

        /// <summary>
        /// Adds schema properties to dictionary.
        /// </summary>
        /// <param name="type">Schema type.</param>
        /// <param name="propDefDictionary">The property definition dictionary.</param>
        internal static void AddSchemaPropertiesToDictionary(
            Type type,
            Dictionary<string, PropertyDefinitionBase> propDefDictionary)
        {
            ForeachPublicStaticPropertyFieldInType(
                type,
                delegate (PropertyDefinition propertyDefinition, FieldInfo fieldInfo)
                {
                    // Some property definitions descend from ServiceObjectPropertyDefinition but don't have
                    // a Uri, like ExtendedProperties. Ignore them.
                    if (!string.IsNullOrEmpty(propertyDefinition.Uri))
                    {
                        if (propDefDictionary.TryGetValue(propertyDefinition.Uri, out PropertyDefinitionBase existingPropertyDefinition))
                        {
                            EwsUtilities.Assert(
                                existingPropertyDefinition == propertyDefinition,
                                "Schema.allSchemaProperties.delegate",
                                string.Format("There are at least two distinct property definitions with the following URI: {0}", propertyDefinition.Uri));
                        }
                        else
                        {
                            propDefDictionary.Add(propertyDefinition.Uri, propertyDefinition);

                            // The following is a "generic hack" to register properties that are not public and
                            // thus not returned by the above GetFields call. It is currently solely used to register
                            // the MeetingTimeZone property.
                            List<PropertyDefinition> associatedInternalProperties = propertyDefinition.GetAssociatedInternalProperties();

                            foreach (PropertyDefinition associatedInternalProperty in associatedInternalProperties)
                            {
                                propDefDictionary.Add(associatedInternalProperty.Uri, associatedInternalProperty);
                            }
                        }
                    }
                });
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ServiceObjectSchema"/> class.
        /// </summary>
        protected ServiceObjectSchema()
        {
            this.RegisterProperties();
        }

        /// <summary>
        /// Finds the property definition.
        /// </summary>
        /// <param name="uri">The URI.</param>
        /// <returns>Property definition.</returns>
        internal static PropertyDefinitionBase FindPropertyDefinition(string uri)
        {
            return allSchemaProperties.Member[uri];
        }

        /// <summary>
        /// Initialize schema property names.
        /// </summary>
        internal static void InitializeSchemaPropertyNames()
        {
            lock (lockObject)
            {
                foreach (Type type in allSchemaTypes.Member)
                {
                    ForeachPublicStaticPropertyFieldInType(
                        type,
                        delegate (PropertyDefinition propDef, FieldInfo fieldInfo) { propDef.Name = fieldInfo.Name; });
                }
            }
        }

        /// <summary>
        /// Defines the ExtendedProperties property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ExtendedProperties =
            new ComplexPropertyDefinition<ExtendedPropertyCollection>(
                XmlElementNames.ExtendedProperty,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.ReuseInstance | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate,
                ExchangeVersion.Exchange2007_SP1,
                delegate () { return new ExtendedPropertyCollection(); });

        private readonly Dictionary<string, PropertyDefinition> properties = new Dictionary<string, PropertyDefinition>();
        private readonly List<PropertyDefinition> visibleProperties = new List<PropertyDefinition>();

        /// <summary>
        /// Registers a schema property.
        /// </summary>
        /// <param name="property">The property to register.</param>
        /// <param name="isInternal">Indicates whether the property is internal or should be visible to developers.</param>
        private void RegisterProperty(PropertyDefinition property, bool isInternal)
        {
            this.properties.Add(property.XmlElementName, property);

            if (!isInternal)
            {
                this.visibleProperties.Add(property);
            }

            // If this property does not have to be requested explicitly, add
            // it to the list of firstClassProperties.
            if (!property.HasFlag(PropertyDefinitionFlags.MustBeExplicitlyLoaded))
            {
                this.FirstClassProperties.Add(property);
            }

            // If this property can be found, add it to the list of firstClassSummaryProperties
            if (property.HasFlag(PropertyDefinitionFlags.CanFind))
            {
                this.FirstClassSummaryProperties.Add(property);
            }
        }

        /// <summary>
        /// Registers a schema property that will be visible to developers.
        /// </summary>
        /// <param name="property">The property to register.</param>
        internal void RegisterProperty(PropertyDefinition property)
        {
            this.RegisterProperty(property, false);
        }

        /// <summary>
        /// Registers an internal schema property.
        /// </summary>
        /// <param name="property">The property to register.</param>
        internal void RegisterInternalProperty(PropertyDefinition property)
        {
            this.RegisterProperty(property, true);
        }

        /// <summary>
        /// Registers an indexed property.
        /// </summary>
        /// <param name="indexedProperty">The indexed property to register.</param>
        internal void RegisterIndexedProperty(IndexedPropertyDefinition indexedProperty)
        {
            this.IndexedProperties.Add(indexedProperty);
        }

        /// <summary>
        /// Registers properties.
        /// </summary>
        internal virtual void RegisterProperties()
        {
        }

        /// <summary>
        /// Gets the list of first class properties for this service object type.
        /// </summary>
        internal List<PropertyDefinition> FirstClassProperties { get; } = new List<PropertyDefinition>();

        /// <summary>
        /// Gets the list of first class summary properties for this service object type.
        /// </summary>
        internal List<PropertyDefinition> FirstClassSummaryProperties { get; } = new List<PropertyDefinition>();

        /// <summary>
        /// Gets the list of indexed properties for this service object type.
        /// </summary>
        internal List<IndexedPropertyDefinition> IndexedProperties { get; } = new List<IndexedPropertyDefinition>();

        /// <summary>
        /// Tries to get property definition.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="propertyDefinition">The property definition.</param>
        /// <returns>True if property definition exists.</returns>
        internal bool TryGetPropertyDefinition(string xmlElementName, out PropertyDefinition propertyDefinition)
        {
            return this.properties.TryGetValue(xmlElementName, out propertyDefinition);
        }

        #region IEnumerable<SimplePropertyDefinition> Members

        /// <summary>
        /// Obtains an enumerator for the properties of the schema.
        /// </summary>
        /// <returns>An IEnumerator instance.</returns>
        public IEnumerator<PropertyDefinition> GetEnumerator()
        {
            return this.visibleProperties.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Obtains an enumerator for the properties of the schema.
        /// </summary>
        /// <returns>An IEnumerator instance.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.visibleProperties.GetEnumerator();
        }

        #endregion
    }
}