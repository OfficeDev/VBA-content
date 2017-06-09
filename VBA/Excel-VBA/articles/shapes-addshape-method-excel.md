---
title: Shapes.AddShape Method (Excel)
keywords: vbaxl10.chm638084
f1_keywords:
- vbaxl10.chm638084
ms.prod: excel
api_name:
- Excel.Shapes.AddShape
ms.assetid: 5d08e6d5-2875-795a-8fe1-f4032d4d3fc0
ms.date: 06/08/2017
---


# Shapes.AddShape Method (Excel)

Returns a  **[Shape](shape-object-excel.md)** object that represents the new AutoShape in a worksheet.


## Syntax

 _expression_ . **AddShape**( **_Type_** , **_Left_** , **_Top_** , **_Width_** , **_Height_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[MsoAutoShapeType](http://msdn.microsoft.com/library/7e6fe414-2b25-56d7-a678-b6e718329118%28Office.15%29.aspx)**|Specifies the type of AutoShape to create.|
| _Left_|Required| **Single**|The position (in points) of the upper-left corner of the AutoShape's bounding box relative to the upper-left corner of the document.|
| _Top_|Required| **Single**|The position (in points) of the upper-left corner of the AutoShape's bounding box relative to the upper-left corner of the document.|
| _Width_|Required| **Single**|The width of the AutoShape's bounding box, in points.|
| _Height_|Required| **Single**|The height of the AutoShape's bounding box, in points.|

### Return Value

Shape


## Remarks

To change the type of an AutoShape that you?ve added, set the  **[AutoShapeType](shape-autoshapetype-property-excel.md)** property.


## Example

This example adds a rectangle to  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddShape msoShapeRectangle, 50, 50, 100, 200
```


## See also


#### Concepts


[Shapes Object](shapes-object-excel.md)

