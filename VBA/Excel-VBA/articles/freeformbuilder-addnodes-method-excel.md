---
title: FreeformBuilder.AddNodes Method (Excel)
keywords: vbaxl10.chm648073
f1_keywords:
- vbaxl10.chm648073
ms.prod: excel
api_name:
- Excel.FreeformBuilder.AddNodes
ms.assetid: 8fff188d-1c47-87f0-8388-2b12534e82c2
ms.date: 06/08/2017
---


# FreeformBuilder.AddNodes Method (Excel)

Adds a point in the current shape and then draws a line from the current node to last node that was added.


## Syntax

 _expression_ . **AddNodes**( **_SegmentType_** , **_EditingType_** , **_X1_** , **_Y1_** , **_X2_** , **_Y2_** , **_X3_** , **_Y3_** )

 _expression_ A variable that represents a **FreeformBuilder** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SegmentType_|Required| **[MsoSegmentType](http://msdn.microsoft.com/library/1a015227-8090-52a7-24f9-71d7e34fd05d%28Office.15%29.aspx)**|The type of segment to be added.|
| _EditingType_|Required| **[MsoEditingType](http://msdn.microsoft.com/library/5fe5c4f6-6467-c6a7-197c-ff700c384b92%28Office.15%29.aspx)**|The editing property of the vertex.|
| _X1_|Required| **Single**|If the  _EditingType_ of the new segment is **msoEditingAuto** , this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the end point of the new segment. If the _EditingType_ of the new node is **msoEditingCorner** , this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the first control point for the new segment.|
| _Y1_|Required| **Single**|If the  _EditingType_ of the new segment is **msoEditingAuto** , this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the end point of the new segment. If the _EditingType_ of the new node is **msoEditingCorner** , this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the first control point for the new segment.|
| _X2_|Optional| **Variant**|If the  _EditingType_ of the new segment is **msoEditingCorner** , this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the _EditingType_ of the new segment is **msoEditingAuto** , don't specify a value for this argument.|
| _Y2_|Optional| **Variant**|If the  _EditingType_ of the new segment is **msoEditingCorner** , this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the _EditingType_ of the new segment is **msoEditingAuto** , don't specify a value for this argument.|
| _X3_|Optional| **Variant**|If the  _EditingType_ of the new segment is **msoEditingCorner** , this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the _EditingType_ of the new segment is **msoEditingAuto** , don't specify a value for this argument.|
| _Y3_|Optional| **Variant**|If the  _EditingType_ of the new segment is **msoEditingCorner** , this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the _EditingType_ of the new segment is **msoEditingAuto** , don't specify a value for this argument.|

## Remarks





| **MsoSegmentType** can be one of these **MsoSegmentType** constants.|
| **msoSegmentLine**|
| **msoSegmentCurve**|


| **MsoEditingType** can be one of these **MsoEditingType** constants.|
| **msoEditingAuto**|
| **msoEditingCorner**|
|Cannot be  **msoEditingSmooth** or **msoEditingSymmetric** If _SegmentType_ is **msoSegmentLine** , _EditingType_ must be **msoEditingAuto** .|

## Example

This example adds a freeform with four segments to  `myDocument`.


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


[FreeformBuilder Object](freeformbuilder-object-excel.md)

