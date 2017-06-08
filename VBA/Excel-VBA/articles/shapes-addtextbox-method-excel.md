---
title: Shapes.AddTextbox Method (Excel)
keywords: vbaxl10.chm638086
f1_keywords:
- vbaxl10.chm638086
ms.prod: excel
api_name:
- Excel.Shapes.AddTextbox
ms.assetid: c594be81-95e6-37da-2c55-418f11ad7554
ms.date: 06/08/2017
---


# Shapes.AddTextbox Method (Excel)

Creates a text box. Returns a  **[Shape](shape-object-excel.md)** object that represents the new text box.


## Syntax

 _expression_ . **AddTextbox**( **_Orientation_** , **_Left_** , **_Top_** , **_Width_** , **_Height_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Orientation_|Required| **[MsoTextOrientation](http://msdn.microsoft.com/library/7e8d0e06-14dd-3ea1-a2e4-50375919517f%28Office.15%29.aspx)**|The orientation of the textbox.|
| _Left_|Required| **Single**|The position (in points) of the upper-left corner of the text box relative to the upper-left corner of the document.|
| _Top_|Required| **Single**|The position (in points) of the upper-left corner of the text box relative to the top of the document.|
| _Width_|Required| **Single**|The width of the text box, in points.|
| _Height_|Required| **Single**|The height of the text box, in points.|

### Return Value

Shape


## Example

This example adds a text box that contains the text "Test Box" to  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, _ 
    100, 100, 200, 50) _ 
    .TextFrame.Characters.Text = "Test Box"
```


## See also


#### Concepts


[Shapes Object](shapes-object-excel.md)

