---
title: ShapeRange.AutoShapeType Property (PowerPoint)
keywords: vbapp10.chm548016
f1_keywords:
- vbapp10.chm548016
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.AutoShapeType
ms.assetid: a1b6c923-dac7-8b5a-6d8b-46a62cfb119e
ms.date: 06/08/2017
---


# ShapeRange.AutoShapeType Property (PowerPoint)

Returns or sets the shape type for the specified  **ShapeRange** object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write.


## Syntax

 _expression_. **AutoShapeType**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

[MsoAutoShapeType](http://msdn.microsoft.com/library/7e6fe414-2b25-56d7-a678-b6e718329118%28Office.15%29.aspx)


## Remarks

Use the  **Type** property of the **[ConnectorFormat](connectorformat-object-powerpoint.md)** object to set or return the connector type.


## Example

This example replaces all 16-point stars with 32-point stars in  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1) 
For Each s In myDocument.Shapes 
    If s.AutoShapeType = msoShape16pointStar Then 
        s.AutoShapeType = msoShape32pointStar 
    End If 
Next
```


 **Note**  When you change the type of a shape, the shape retains its size, color, and other attributes.


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

