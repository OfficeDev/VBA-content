---
title: ShapeRange.Title Property (PowerPoint)
keywords: vbapp10.chm548097
f1_keywords:
- vbapp10.chm548097
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Title
ms.assetid: bb4e08a3-6517-c500-23ac-ec65b3340f76
ms.date: 06/08/2017
---


# ShapeRange.Title Property (PowerPoint)

Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the slide title. Read-only.


## Syntax

 _expression_. **Title**

 _expression_ A variable that represents a **ShapeRange** object.


## Remarks

You can also use the  **[Item](placeholders-item-method-powerpoint.md)** method of the **[Shapes](shapes-object-powerpoint.md)** or **[Placeholders](placeholders-object-powerpoint.md)** collection to return the slide title.


## Example

The following example sets the title text on  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.Title.TextFrame.TextRange.Text = "Welcome!"
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

