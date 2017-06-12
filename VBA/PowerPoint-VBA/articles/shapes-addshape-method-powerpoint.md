---
title: Shapes.AddShape Method (PowerPoint)
keywords: vbapp10.chm543012
f1_keywords:
- vbapp10.chm543012
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddShape
ms.assetid: 2bc6cce5-3461-61ff-083d-bd36ee71cb59
ms.date: 06/08/2017
---


# Shapes.AddShape Method (PowerPoint)

Creates an AutoShape. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new AutoShape.


## Syntax

 _expression_. **AddShape**( **_Type_**, **_Left_**, **_Top_**, **_Width_**, **_Height_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**[MsoAutoShapeType](http://msdn.microsoft.com/library/7e6fe414-2b25-56d7-a678-b6e718329118%28Office.15%29.aspx)**|Specifies the type of AutoShape to create.|
| _Left_|Required|**Single**|The position, measured in points, of the left edge of the AutoShape relative to the left edge of the slide.|
| _Top_|Required|**Single**|The position, measured in points, of the top edge of the AutoShape relative to the top edge of the slide.|
| _Width_|Required|**Single**|The width of the AutoShape, measured in points.|
| _Height_|Required|**Single**|The height of the AutoShape, measured in points.|

### Return Value

Shape


## Remarks

To change the type of an AutoShape that you've added, set the  **AutoShapeType** property.


## Example

This example adds a rectangle to  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1) 
myDocument.Shapes.AddShape Type:=msoShapeRectangle, _ 
    Left:=50, Top:=50, Width:=100, Height:=200
```


## See also


#### Concepts


[Shapes Object](shapes-object-powerpoint.md)

