---
title: Shape.Visible Property (Word)
keywords: vbawd10.chm161480831
f1_keywords:
- vbawd10.chm161480831
ms.prod: word
api_name:
- Word.Shape.Visible
ms.assetid: b3024bf2-3015-d3ce-97dc-2dd5858bf798
ms.date: 06/08/2017
---


# Shape.Visible Property (Word)

 **True** if the specified object, or the formatting applied to it, is visible. Read/write **MsoTriState** .


## Syntax

 _expression_ . **Visible**

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

FSome methods and properties may be unavailable if the  **Visible** property is **False** .


## Example

This example creates a new document and then adds text and a rectangle to it. The example also sets Word to hide the rectangle while the document is being printed and then to make it visible again after printing is completed.


```vb
Set myDoc = Documents.Add 
Selection.TypeText Text:="This is some sample text." 
With myDoc 
 .Shapes.AddShape msoShapeRectangle, 200, 70, 150, 60 
 .Shapes(1).Visible = False 
 .PrintOut 
 .Shapes(1).Visible = True 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

