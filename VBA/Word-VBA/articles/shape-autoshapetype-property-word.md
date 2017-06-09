---
title: Shape.AutoShapeType Property (Word)
keywords: vbawd10.chm161480805
f1_keywords:
- vbawd10.chm161480805
ms.prod: word
api_name:
- Word.Shape.AutoShapeType
ms.assetid: 521ed05e-99b5-d917-6a26-3d911192b569
ms.date: 06/08/2017
---


# Shape.AutoShapeType Property (Word)

Returns or sets the shape type for the specified  **Shape** object, which must represent an AutoShape other than a line or freeform drawing. Read/write **MsoAutoShapeType** .


## Syntax

 _expression_ . **AutoShapeType**

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

When you change the type of a shape, the shape retains its size, color, and other attributes.


## Example

This example replaces all 16-point stars with 32-point stars in the active document.


```vb
Sub ReplaceAutoShape() 
 Dim docNew As Document 
 Dim shpStar As Shape 
 Set docNew = ActiveDocument 
 For Each shpStar In docNew.Shapes 
 If shpStar.AutoShapeType = msoShape16pointStar Then 
 shpStar.AutoShapeType = msoShape32pointStar 
 End If 
 Next 
End Sub
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

