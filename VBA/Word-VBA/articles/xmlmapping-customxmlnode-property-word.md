---
title: XMLMapping.CustomXMLNode Property (Word)
keywords: vbawd10.chm199688194
f1_keywords:
- vbawd10.chm199688194
ms.prod: word
api_name:
- Word.XMLMapping.CustomXMLNode
ms.assetid: c28e3a1e-1bc3-fbe7-7ff8-78adef326bbd
ms.date: 06/08/2017
---


# XMLMapping.CustomXMLNode Property (Word)

Returns a  **CustomXMLNode** object that represents the custom XML node in the data store to which the content control in the document maps.


## Syntax

 _expression_ . **CustomXMLNode**

 _expression_ An expression that returns an **[XMLMapping](xmlmapping-object-word.md)** object.


## Example

The following example inserts a new content control and custom XML part into the active document, maps the content control to a node in the custom XML part, and then sets the value of the mapped XML node.


```vb
Dim objCC As ContentControl 
Dim objPart As CustomXMLPart 
Dim objNode As CustomXMLNode 
Dim objMap As XMLMapping 
 
Set objCC = ActiveDocument.ContentControls.Add(wdContentControlText) 
Set objPart = ActiveDocument.CustomXMLParts.Add("<books><book>" &; _ 
 "<author></author><title></title><genre></genre><price></price>" &; _ 
 "<pub_date></pub_date><abstract></abstract></book></books>") 
 
Set objMap = objCC.XMLMapping 
objMap.SetMapping "/books/book/author", , objPart 
 
Set objNode = objMap.CustomXMLNode 
objNode.Text = "Matt Hink" 
 
objCC.Range.Text = objNode.Text
```


## See also


#### Concepts


[XMLMapping Object](xmlmapping-object-word.md)

