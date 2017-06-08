---
title: ShapeRange.AutoShapeType Property (Excel)
keywords: vbaxl10.chm640098
f1_keywords:
- vbaxl10.chm640098
ms.prod: excel
api_name:
- Excel.ShapeRange.AutoShapeType
ms.assetid: de4c8273-2804-012c-38a0-7689aa54b02e
ms.date: 06/08/2017
---


# ShapeRange.AutoShapeType Property (Excel)

Returns or sets the shape type for the specified  **[Shape](shape-object-excel.md)** or **[ShapeRange](shaperange-object-excel.md)** object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write **[MsoAutoShapeType](http://msdn.microsoft.com/library/7e6fe414-2b25-56d7-a678-b6e718329118%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **AutoShapeType**

 _expression_ A variable that represents a **ShapeRange** object.


## Remarks

When you change the type of a shape, the shape retains its size, color, and other attributes.

Use the  **[Type](connectorformat-type-property-excel.md)** property of the **[ConnectorFormat](connectorformat-object-excel.md)** object to set or return the connector type.


## Example

This example replaces all 16-point stars with 32-point stars in  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
    If s.AutoShapeType = msoShape16pointStar Then 
        s.AutoShapeType = msoShape32pointStar 
    End If 
Next
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

