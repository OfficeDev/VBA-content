---
title: Rectangle.RectangleType Property (Word)
keywords: vbawd10.chm234029058
f1_keywords:
- vbawd10.chm234029058
ms.prod: word
api_name:
- Word.Rectangle.RectangleType
ms.assetid: 0d483c0d-5a97-7f68-d7fa-7458063b6643
ms.date: 06/08/2017
---


# Rectangle.RectangleType Property (Word)

Returns a  **WdRectangleType** constant that represents the type for the specified rectangle.


## Syntax

 _expression_ . **RectangleType**

 _expression_ Required. A variable that represents a **[Rectangle](rectangle-object-word.md)** object.


## Remarks

Rectangles in Microsoft Word are sections within a page in a document that contain specific types of information. Some sections are portions of text; others are shapes. The purpose of rectangles is to allow more control over programmatic page layout.


## Example

The following example accesses the first rectangle on the first page in the active document, and if it is a text rectangle, checks the spelling.


```vb
Dim objRectangle As Rectangle 
 
Set objRectangle = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages(1).Rectangles(1) 
 
If objRectangle.RectangleType = wdTextRectangle Then 
 objRectangle.Range.CheckSpelling 
End If
```


## See also


#### Concepts


[Rectangle Object](rectangle-object-word.md)

