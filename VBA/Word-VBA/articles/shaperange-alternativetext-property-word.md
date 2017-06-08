---
title: ShapeRange.AlternativeText Property (Word)
keywords: vbawd10.chm162857091
f1_keywords:
- vbawd10.chm162857091
ms.prod: word
api_name:
- Word.ShapeRange.AlternativeText
ms.assetid: c8d98235-942e-7c1f-bd00-5a378c060ec7
ms.date: 06/08/2017
---


# ShapeRange.AlternativeText Property (Word)

Returns or sets the alternative text associated with a shape in a Web page. Read/write  **String** .


## Syntax

 _expression_ . **AlternativeText**

 _expression_ A variable that represents a **[ShapeRange](shaperange-object-word.md)** object.


## Example

The following example sets the alternative text for the selected shape in the active window. The selected shape is a picture of a mallard duck.


```vb
ActiveWindow.Selection.ShapeRange _ 
 .AlternativeText = "This is a mallard duck."
```


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

