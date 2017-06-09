---
title: Shapes.AddLabel Method (Excel)
keywords: vbaxl10.chm638080
f1_keywords:
- vbaxl10.chm638080
ms.prod: excel
api_name:
- Excel.Shapes.AddLabel
ms.assetid: eb0bfb2a-51ab-ce65-0ef2-aa964d3b08ba
ms.date: 06/08/2017
---


# Shapes.AddLabel Method (Excel)

Creates a label. Returns a  **[Shape](shape-object-excel.md)** object that represents the new label.


## Syntax

 _expression_ . **AddLabel**( **_Orientation_** , **_Left_** , **_Top_** , **_Width_** , **_Height_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Orientation_|Required| **[MsoTextOrientation](http://msdn.microsoft.com/library/7e8d0e06-14dd-3ea1-a2e4-50375919517f%28Office.15%29.aspx)**|The text orientation within the label.|
| _Left_|Required| **Single**|The position (in points) of the upper-left corner of the label relative to the upper-left corner of the document.|
| _Top_|Required| **Single**|The position (in points) of the upper-left corner of the label relative to the top corner of the document.|
| _Width_|Required| **Single**|The width of the label, in points.|
| _Height_|Required| **Single**|The height of the label, in points.|

### Return Value

Shape


## Example

This example adds a vertical label that contains the text "Test Label" to  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddLabel(msoTextOrientationVertical, _ 
    100, 100, 60, 150) _ 
    .TextFrame.Characters.Text = "Test Label"
```


## See also


#### Concepts


[Shapes Object](shapes-object-excel.md)

