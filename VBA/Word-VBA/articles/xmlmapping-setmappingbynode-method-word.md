---
title: XMLMapping.SetMappingByNode Method (Word)
keywords: vbawd10.chm199688197
f1_keywords:
- vbawd10.chm199688197
ms.prod: word
api_name:
- Word.XMLMapping.SetMappingByNode
ms.assetid: 8eab3471-e1dc-f7ec-9b45-9fb459088190
ms.date: 06/08/2017
---


# XMLMapping.SetMappingByNode Method (Word)

Allows creating or changing the XML data mapping on a content control. Returns  **True** if Microsoft Word maps the content control to a custom XML node in the document?s custom XML data store.


## Syntax

 _expression_ . **SetMappingByNode**( **_Node_** )

 _expression_ An expression that returns an **[XMLMapping](xmlmapping-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Node_|Required| **CustomXMLNode**|Specifies the XML node to which to map the current content control.|

### Return Value

Boolean


## Remarks

If the XML mapping already exists, then Word replaces the existing XML mapping, and the contents of the new mapped XML node replaces the text of the content control. See also the  **[SetMapping](xmlmapping-setmapping-method-word.md)** method.


 **Note**  Creating a mapping for a rich-text content control causes a run-time error.


## Example

The following example sets the built-in document property for the document author, inserts a new content control into the active document, and then sets the XML mapping for the control to the built-in document property.


```vb
Dim objcc As ContentControl 
Dim objNode As CustomXMLNode 
Dim objMap As XMLMapping 
Dim blnMap As Boolean 
 
ActiveDocument.BuiltInDocumentProperties("Author").Value = "David Jaffe" 
 
Set objcc = ActiveDocument.ContentControls.Add _ 
 (wdContentControlDate, ActiveDocument.Paragraphs(1).Range) 
 
Set objNode = ActiveDocument.CustomXMLParts.SelectByNamespace _ 
 ("http://schemas.openxmlformats.org/package/2006/metadata/core-properties") _ 
 (1).DocumentElement.ChildNodes(1) 
 
Set objMap = objcc.XMLMapping 
blnMap = objMap.SetMappingByNode(objNode)
```

The following example creates a custom XML part, and then creates two content controls and maps each content control to a specific node within the custom XML.




```vb
Dim objRange As Range 
Dim objCustomPart As CustomXMLPart 
Dim objCustomControl As ContentControl 
Dim objCustomNode As CustomXMLNode 
 
Set objCustomPart = ActiveDocument.CustomXMLParts.Add 
objCustomPart.LoadXML ("<books><book><author>Matt Hink</author>" &; _ 
 "<title>Migration Paths of the Red Breasted Robin</title><genre>non-fiction</genre>" &; _ 
 "<price>29.95</price><pub_date>2007-02-01</pub_date><abstract>" &; _ 
 "You see them in the spring outside your windows. You hear their lovely " &; _ 
 "songs wafting in the warm spring air. Now follow the path of the red breasted robin " &; _ 
 "as it migrates to warmer climes in the fall, and then back to your back yard " &; _ 
 "in the spring.</abstract></book></books>") 
 
ActiveDocument.Range.InsertParagraphBefore 
Set objRange = ActiveDocument.Paragraphs(1).Range 
Set objCustomNode = objCustomPart.SelectSingleNode _ 
 ("/books/book/title") 
Set objCustomControl = ActiveDocument.ContentControls _ 
 .Add(wdContentControlText, objRange) 
objCustomControl.XMLMapping.SetMappingByNode objCustomNode 
 
objRange.InsertParagraphAfter 
Set objRange = ActiveDocument.Paragraphs(2).Range 
Set objCustomNode = objCustomPart.SelectSingleNode _ 
 ("/books/book/abstract") 
Set objCustomControl = ActiveDocument.ContentControls _ 
 .Add(wdContentControlText, objRange) 
objCustomControl.XMLMapping.SetMappingByNode objCustomNode
```


## See also


#### Concepts


[XMLMapping Object](xmlmapping-object-word.md)

