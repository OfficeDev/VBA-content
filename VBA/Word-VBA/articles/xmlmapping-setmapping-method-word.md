---
title: XMLMapping.SetMapping Method (Word)
keywords: vbawd10.chm199688195
f1_keywords:
- vbawd10.chm199688195
ms.prod: word
api_name:
- Word.XMLMapping.SetMapping
ms.assetid: 0d33be39-f355-7a59-802c-33d031485a0e
ms.date: 06/08/2017
---


# XMLMapping.SetMapping Method (Word)

Allows creating or changing the XML mapping on a content control. Returns  **True** if Microsoft Word maps the content control to a custom XML node in the document?s custom XML data store.


## Syntax

 _expression_ . **SetMapping**( **_XPath_** , **_PrefixMapping_** , **_Source_** )

 _expression_ An expression that returns an **[XMLMapping](xmlmapping-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XPath_|Required| **String**|Specifies an XPath string that represents the XML node to which to map the content control. An invalid XPath string causes a run-time error.|
| _PrefixMapping_|Optional| **String**|Specifies the prefix mappings to use when querying the expression provided in the XPath parameter. If omitted, Word uses the set of prefix mappings for the specified custom XML part in the current document.|
| _Source_|Optional| **CustomXMLPart**|Specifies the desired custom XML data to which to map the content control. If this parameter is omitted, the XPath is evaluated against all custom XML in the current document, and the mapping is established with the first  **CustomXMLPart** in which the XPath resolves to an XML node.|

### Return Value

Boolean


## Remarks

If the XML mapping already exists, Word replaces the existing XML mapping and the contents of the new mapped XML node replaces the text of the content control. If the specified XPath does not evaluate to an XML node in the specified custom XML part or parts, you can still specify the mapping, and one will be created. This mapping automatically links when the specified XPath would evaluate to an XML node in the specified custom XML parts.

See also the  **[SetMappingByNode](xmlmapping-setmappingbynode-method-word.md)** method.


 **Note**  Creating a mapping for a rich-text content control causes a run-time error.


## Example

The following example inserts a custom XML part and sets the XML for the custom part, and then inserts two content controls at the beginning of the document and maps the contents of the controls to the contents of XML elements in the custom part.


```vb
Dim objRange As Range 
Dim objCustomPart As CustomXMLPart 
Dim objCustomControl As ContentControl 
 
Set objCustomPart = ActiveDocument.CustomXMLParts.Add 
objCustomPart.LoadXML ("<books><book><author>Matt Hink</author>" &; _ 
 "<title>Migration Paths of the Red Breasted Robin</title>" &; _ 
 "<genre>non-fiction</genre><price>29.95</price>" &; _ 
 "<pub_date>2/1/2007</pub_date><abstract>You see them in " &; _ 
 "the spring outside your windows. You hear their lovely " &; _ 
 "songs wafting in the warm spring air. Now follow the path " &; _ 
 "of the red breasted robin as it migrates to warmer climes " &; _ 
 "in the fall, and then back to your back yard in the spring." &; _ 
 "</abstract></book></books>") 
 
ActiveDocument.Range.InsertParagraphBefore 
Set objRange = ActiveDocument.Paragraphs(1).Range 
Set objCustomControl = ActiveDocument.ContentControls _ 
 .Add(wdContentControlText, objRange) 
objCustomControl.XMLMapping.SetMapping _ 
 "/books/book/title", , objCustomPart 
 
objRange.InsertParagraphAfter 
Set objRange = ActiveDocument.Paragraphs(2).Range 
Set objCustomControl = ActiveDocument.ContentControls _ 
 .Add(wdContentControlText, objRange) 
objCustomControl.XMLMapping.SetMapping _ 
 "/books/book/abstract", , objCustomPart
```


## See also


#### Concepts


[XMLMapping Object](xmlmapping-object-word.md)

