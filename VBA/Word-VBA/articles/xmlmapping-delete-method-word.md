---
title: XMLMapping.Delete Method (Word)
keywords: vbawd10.chm199688196
f1_keywords:
- vbawd10.chm199688196
ms.prod: word
api_name:
- Word.XMLMapping.Delete
ms.assetid: 72864b8d-5b59-66c3-b9e3-b970f8adf7aa
ms.date: 06/08/2017
---


# XMLMapping.Delete Method (Word)

Deletes the XML mapping from the parent content control.


## Syntax

 _expression_ . **Delete**

 _expression_ An expression that returns an **[XMLMapping](xmlmapping-object-word.md)** object.


## Remarks

This operation removes the XML mapping. Both the XML data and the content control remain in the document.


## Example

The following example deletes the XML mapping for all content controls in the active document that are currently mapped.


```vb
Dim objCC As ContentControl 
 
For Each objCC In ActiveDocument.ContentControls 
 If objCC.XMLMapping.IsMapped Then 
 objCC.XMLMapping.Delete 
 End If 
Next
```


## See also


#### Concepts


[XMLMapping Object](xmlmapping-object-word.md)

