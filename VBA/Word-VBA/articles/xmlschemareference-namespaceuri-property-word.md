---
title: XMLSchemaReference.NamespaceURI Property (Word)
keywords: vbawd10.chm32505858
f1_keywords:
- vbawd10.chm32505858
ms.prod: word
api_name:
- Word.XMLSchemaReference.NamespaceURI
ms.assetid: 4081b67e-45d9-13f4-4faa-bcd92c2533b6
ms.date: 06/08/2017
---


# XMLSchemaReference.NamespaceURI Property (Word)

Returns a  **String** that represents the Uniform Resource Identifier (URI) of the schema namespace for the specified object. Read-only.


## Syntax

 _expression_ . **NamespaceURI**

 _expression_ An expression that returns a **XMLSchemaReference** object.


## Remarks

If you are authoring XML schemas for use with Microsoft Word, it is highly recommended that you specify the targetNamespace setting in the schema.


## Example

The following example reloads the SimpleSample schema or, if the schema is not attached to the active document, attaches it.


 **Note**  The SimpleSample schema is included in the Smart Document Software Development Kit (SDK). For more information, refer to the Smart Document SDK on the Microsoft Developer Network (MSDN) Web site.


```vb
If ActiveDocument.XMLSchemaReferences.Item(1) _ 
 .NamespaceURI <> "SimpleSample" Then 
 
 Application.XMLNamespaces.Item("SimpleSample") _ 
 .AttachToDocument (ActiveDocument) 
 
End If
```


## See also


#### Concepts


[XMLSchemaReference Object](xmlschemareference-object-word.md)

