---
title: CustomXMLPart Members (Office)
ms.prod: office
ms.assetid: 76fe85f4-5a35-7d12-2989-6f17a094dcdf
ms.date: 06/08/2017
---


# CustomXMLPart Members (Office)
Represents a single  **CustomXMLPart** in a **CustomXMLParts** collection.

Represents a single  **CustomXMLPart** in a **CustomXMLParts** collection.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[NodeAfterDelete](customxmlpart-nodeafterdelete-event-office.md)|Occurs after a node is deleted in a  **CustomXMLPart** object.|
|[NodeAfterInsert](customxmlpart-nodeafterinsert-event-office.md)|Occurs after a node is inserted in a  **CustomXMLPart** object.|
|[NodeAfterReplace](customxmlpart-nodeafterreplace-event-office.md)|Occurs just after a node is replaced in a  **CustomXMLPart** object.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddNode](customxmlpart-addnode-method-office.md)|Adds a node to the XML tree.|
|[Delete](customxmlpart-delete-method-office.md)|Deletes the current  **CustomXMLPart** from the data store ( **IXMLDataStore** interface).|
|[Load](customxmlpart-load-method-office.md)|Allows the template author to populate a  **CustomXMLPart** from an existing file. Returns **True** if the load was successful.|
|[LoadXML](customxmlpart-loadxml-method-office.md)|Allows the template author to populate a  **CustomXMLPart** object from an XML string. Returns **True** if the load was successful.|
|[SelectNodes](customxmlpart-selectnodes-method-office.md)|Selects a collection of nodes from a custom XML part.|
|[SelectSingleNode](customxmlpart-selectsinglenode-method-office.md)|Selects a single node within a custom XML part matching an XPath expression.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](customxmlpart-application-property-office.md)|Gets an  **Application** object that represents the container application for the **CustomXMLPart** object. Read-only.|
|[BuiltIn](customxmlpart-builtin-property-office.md)|Gets a value that indicates whether the  **CustomXMLPart** is built-in. Read-only|
|[Creator](customxmlpart-creator-property-office.md)|Gets a 32-bit integer that indicates the application in which the  **CustomXMLPart** object was created. Read-only.|
|[DocumentElement](customxmlpart-documentelement-property-office.md)|Gets the root element of a bound region of data in a document. If the region is empty, the property returns  **Nothing**. Read-only.|
|[Errors](customxmlpart-errors-property-office.md)|Gets a  **CustomXMLValidationErrors** object that provides access to any XML validation errors, if any exists. If no validation errors exist, this property returns **Nothing**. Read-only.|
|[Id](customxmlpart-id-property-office.md)|Gets a  **String** containing the GUID assigned to the current **CustomXMLPart** object. Read-only.|
|[NamespaceManager](customxmlpart-namespacemanager-property-office.md)|Gets the set of namespace prefix mappings used against the current  **CustomXMLPart** object. Read-only.|
|[NamespaceURI](customxmlpart-namespaceuri-property-office.md)|Gets the unique address identifier for the namespace of the  **CustomXMLPart** object. Read-only.|
|[Parent](customxmlpart-parent-property-office.md)|Gets the  **Parent** object for the **CustomXMLPart** object. Read-only.|
|[SchemaCollection](customxmlpart-schemacollection-property-office.md)|Gets or sets a  **CustomXMLSchemaCollection** object representing the set of schemas attached to a bound region of data in a document. Read/write.|
|[XML](customxmlpart-xml-property-office.md)|Gets the XML representation of the current  **CustomXMLPart** object. Read-only.|

