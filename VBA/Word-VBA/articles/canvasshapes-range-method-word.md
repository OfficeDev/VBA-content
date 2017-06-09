---
title: CanvasShapes.Range Method (Word)
keywords: vbawd10.chm7536661
f1_keywords:
- vbawd10.chm7536661
ms.prod: word
api_name:
- Word.CanvasShapes.Range
ms.assetid: 4e0e24aa-7510-a002-63d2-e6dbb5bc3398
ms.date: 06/08/2017
---


# CanvasShapes.Range Method (Word)

Returns a  **ShapeRange** object.


## Syntax

 _expression_ . **Range**( **_Index_** )

 _expression_ Required. A variable that represents a **[CanvasShapes](canvasshapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|Specifies which shapes are to be included in the specified range. Can be an integer that specifies the index number of a shape within the  **Shapes** collection, a string that specifies the name of a shape, or a array that contains integers or strings.|

### Return Value

ShapeRange


## Remarks

Character position values begin with 0 (zero) at the beginning of the document. All characters are counted, including nonprinting characters. Hidden characters are counted even if they're not displayed.

 **ShapeRange** objects don't include **InlineShape** objects. An **InlineShape** object is equivalent to a character and is positioned as a character within a range of text. **Shape** objects are anchored to a range of text (the selection, by default), but they can be positioned anywhere on the page. A **Shape** object will always appear on the same page as the range it is anchored to.

Most operations that you can do with a  **Shape** object you can also do with a **ShapeRange** object that contains a single shape. Some operations, when performed on a **ShapeRange** object that contains multiple shapes, produce an error.


## Example

This example selects and deletes the shapes in the first shape in the active document. This example assumes that the first shape is a canvas shape.


```vb
Sub CanvasShapeRange() 
 Dim rngCanvasShapes As Range 
 Set rngCanvasShapes = ActiveDocument.Shapes(1).CanvasItems.Range(1) 
 rngCanvasShapes.Select 
 rngCanvasShapes.Delete 
End Sub
```


## See also


#### Concepts


[CanvasShapes Collection](canvasshapes-object-word.md)

