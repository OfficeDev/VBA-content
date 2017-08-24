---
title: InlineShapes.Item Method (Publisher)
keywords: vbapb10.chm5767168
f1_keywords:
- vbapb10.chm5767168
ms.prod: publisher
api_name:
- Publisher.InlineShapes.Item
ms.assetid: 7cc4bb2a-e7d8-68c1-7d09-9b81a9d6b87a
ms.date: 06/08/2017
---


# InlineShapes.Item Method (Publisher)

Returns a  **[Shape](shape-object-publisher.md)** object that represents an inline shape contained in a text range. This method is the default member of the **InlineShapes** collection.


## Syntax

 _expression_. **Item**( **_Index_**)

 _expression_A variable that represents an  **InlineShapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|var|Required| **Variant**|The index position or name of the object to return. If  **Index** is an integer, the index into the collection is 1-based. If **Index** is a string, the name of the shape is used as the index. An automation error is returned if the index or name does not represent a shape in the collection.|

### Return Value

Shape


## Example

This example finds the first inline shape in a text range and flips it vertically.


```vb
Dim theShape As Shape 
 
Set theShape = ActiveDocument.Pages(1).Shapes(1) 
 
With theShape.TextFrame.Story.TextRange 
 With .InlineShapes.Item(1) 
 .Flip (msoFlipVertical) 
 End With 
End With
```


