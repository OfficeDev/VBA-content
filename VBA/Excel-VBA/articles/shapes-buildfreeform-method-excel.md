---
title: Shapes.BuildFreeform Method (Excel)
keywords: vbaxl10.chm638087
f1_keywords:
- vbaxl10.chm638087
ms.prod: excel
api_name:
- Excel.Shapes.BuildFreeform
ms.assetid: 0eec4b87-1a40-1e60-a66a-a8bb2b2f7efa
ms.date: 06/08/2017
---


# Shapes.BuildFreeform Method (Excel)

Builds a freeform object. Returns a  **[FreeformBuilder](freeformbuilder-object-excel.md)** object that represents the freeform as it is being built. Use the **[AddNodes](freeformbuilder-addnodes-method-excel.md)** method to add segments to the freeform. After you have added at least one segment to the freeform, you can use the **[ConvertToShape](freeformbuilder-converttoshape-method-excel.md)** method to convert the **FreeformBuilder** object into a **[Shape](shape-object-excel.md)** object that has the geometric description you?ve defined in the **FreeformBuilder** object.


## Syntax

 _expression_ . **BuildFreeform**( **_EditingType_** , **_X1_** , **_Y1_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EditingType_|Required| **[MsoEditingType](http://msdn.microsoft.com/library/5fe5c4f6-6467-c6a7-197c-ff700c384b92%28Office.15%29.aspx)**|The editing property of the first node.|
| _X1_|Required| **Single**|The position (in points) of the first node in the freeform drawing relative to the upper-left corner of the document.|
| _Y1_|Required| **Single**|The position (in points) of the first node in the freeform drawing relative to the upper-left corner of the document.|

### Return Value

FreeformBuilder


## Example

This example adds a freeform with five vertices to  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200) 
    .AddNodes msoSegmentCurve, msoEditingCorner, _ 
        380, 230, 400, 250, 450, 300 
    .AddNodes msoSegmentCurve, msoEditingAuto, 480, 200 
    .AddNodes msoSegmentLine, msoEditingAuto, 480, 400 
    .AddNodes msoSegmentLine, msoEditingAuto, 360, 200 
    .ConvertToShape 
End With
```


## See also


#### Concepts


[Shapes Object](shapes-object-excel.md)

