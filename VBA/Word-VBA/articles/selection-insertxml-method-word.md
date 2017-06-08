---
title: Selection.InsertXML Method (Word)
keywords: vbawd10.chm158663684
f1_keywords:
- vbawd10.chm158663684
ms.prod: word
api_name:
- Word.Selection.InsertXML
ms.assetid: 7a9e52b5-9b05-f939-6fd0-33a923989f48
ms.date: 06/08/2017
---


# Selection.InsertXML Method (Word)

Inserts the specified XML into the document at the cursor, replacing any selected text.


## Syntax

 _expression_ . **InsertXML**( **_XML_** , **_Transform_** )

 _expression_ An expression that returns a **Selection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XML_|Required| **String**|Specifies the XML to insert. This can be any valid custom XML.|
| _Transform_|Optional| **Variant**|Specifies the XML Transformation (XSLT) used to transform the XML. If omitted, the XML is inserted as custom XML without applying a transform.|

### Return Value

Nothing


## Example

The following example inserts the specified XML string into the document at the cursor, or replacing the active selection.


```vb
Dim strXML As String 
 
strXML = "<?xml version=""1.0""?><abc:books xmlns:abc=""urn:books"" " &; _ 
 "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " &; _ 
 "xsi:schemaLocation=""urn:books books.xsd""><book>" &; _ 
 "<author>Matt Hink</author><title>Migration Paths of the Red " &; _ 
 "Breasted Robin</title><genre>non-fiction</genre>" &; _ 
 "<price>29.95</price><pub_date>2006-05-01</pub_date>" &; _ 
 "<abstract>You see them in the spring outside your windows. " &; _ 
 "You hear their lovely songs wafting in the warm spring air. " &; _ 
 "Now follow their path as they migrate to warmer climes in the fall, " &; _ 
 "and then back to your back yard in the spring.</abstract></book></abc:books>" 
 
Selection.InsertXML strXML
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

