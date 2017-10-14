---
title: InlineShapes.Range Property (Publisher)
keywords: vbapb10.chm5767173
f1_keywords:
- vbapb10.chm5767173
ms.prod: publisher
api_name:
- Publisher.InlineShapes.Range
ms.assetid: 375843c1-5198-6981-2e7c-8abd1d0e9dff
ms.date: 06/08/2017
---


# InlineShapes.Range Property (Publisher)

Returns a  **[ShapeRange](shaperange-object-publisher.md)** collection that represents the same set of inline shapes as the **InlineShapes** collection whose method was called. This allows for miscellaneous formatting of the contained shapes. Read-only.


## Syntax

 _expression_. **Range**( **_Index_**)

 _expression_A variable that represents an  **InlineShapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Optional| **Long**|The index position of the inline shape within the  **ShapeRange** collection.|

## Example

The following example searches through each shape on the first page of the publication, and for all inline shapes within each shape, finds the first inline shape within the range of inline shapes and flips it vertically.


```vb
Dim theShape As Shape 
Dim theShapes As Shapes 
 
Set theShapes = ActiveDocument.Pages(1).Shapes 
 
For Each theShape In theShapes 
 With theShape.TextFrame.TextRange 
 .InlineShapes.Range(1).Flip (msoFlipVertical) 
 End With 
Next
```


