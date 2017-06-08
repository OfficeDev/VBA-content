---
title: ShapeRange.Title Property (Word)
keywords: vbawd10.chm162857166
f1_keywords:
- vbawd10.chm162857166
ms.prod: word
api_name:
- Word.ShapeRange.Title
ms.assetid: 0ee66d0b-9c32-7975-2e50-3205a15730a5
ms.date: 06/08/2017
---


# ShapeRange.Title Property (Word)

Returns or sets a  **String** that contains a title for the shapes in the specified shape range. Read/write.


## Syntax

 _expression_ . **Title**

 _expression_ A variable that represents a **[ShapeRange](shaperange-object-word.md)** object.


## Remarks

Use the  **Title** property to provide an alternative text title for a shape. This property adds title text to the **Title** text box on the **Alt Text** pane of the **Format Shape** dialog in Word.


 **Note**  Web browsers display alternative text while tables are loading or if they are missing. Web search engines use the alternative text to help find Web pages. Alternative text is also used to assist disabilities.


## Example

The following code example adds an alternative text title to the first and third shapes in the active document.


```vb
ActiveDocument.Shapes.Range(Array(1, 3)).Title = "Part of a shape array."
```


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

