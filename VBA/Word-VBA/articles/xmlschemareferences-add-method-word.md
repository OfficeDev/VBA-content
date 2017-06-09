---
title: XMLSchemaReferences.Add Method (Word)
keywords: vbawd10.chm116129893
f1_keywords:
- vbawd10.chm116129893
ms.prod: word
api_name:
- Word.XMLSchemaReferences.Add
ms.assetid: fe6fa7d5-287a-f79f-b83b-fc182002504a
ms.date: 06/08/2017
---


# XMLSchemaReferences.Add Method (Word)

Returns an  **XMLSchemaReference** that represents a schema applied to a document.


## Syntax

 _expression_ . **Add**( **_NamespaceURI_** , **_Alias_** , **_FileName_** , **_InstallForAllUsers_** )

 _expression_ Required. A variable that represents a **[XMLSchemaReferences](xmlschemareferences-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NamespaceURI_|Optional| **String**|The name of the schema as defined in the schema. The Namespace parameter is case-sensitive and must be spelled exactly as it appears in the schema. If the specified namespace cannot be found in any of the schemas attached to the document, an error is displayed.|
| _Alias_|Optional| **String**|The name of the schema as it appears on the  **Schemas** tab in the **Templates and Add-ins** dialog box.|
| _FileName_|Optional| **String**|The path and file name of the schema. This may be a local file path, a network path, or an Internet address.|
| _InstallForAllUsers_|Optional| **Boolean**| **True** if all users that log on to a computer can access and use the new schema. The default is **False** .|

### Return Value

XMLSchemaReference


## Example

The following example attaches the specified schema to the active document. This example assumes that you have an xsd file located at the path specified in the FileName parameter.


```vb
Sub AddSchema() 
 Dim objSchema As XMLSchemaReference 
 
 Set objSchema = ActiveDocument.XMLSchemaReferences _ 
 .Add(, , "c:\schemas\simplesample.xsd", True) 
End Sub
```


## See also


#### Concepts


[XMLSchemaReferences Collection](xmlschemareferences-object-word.md)

