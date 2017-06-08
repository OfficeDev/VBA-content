---
title: ShapeNodes.SetSegmentType Method (Excel)
keywords: vbaxl10.chm112011
f1_keywords:
- vbaxl10.chm112011
ms.prod: excel
api_name:
- Excel.ShapeNodes.SetSegmentType
ms.assetid: 6223e503-4838-2365-9610-26d0a376ccae
ms.date: 06/08/2017
---


# ShapeNodes.SetSegmentType Method (Excel)

Sets the segment type of the segment that follows the node specified by  _Index_. If the node is a control point for a curved segment, this method sets the segment type for that curve. Note that this may affect the total number of nodes by inserting or deleting adjacent nodes.


## Syntax

 _expression_ . **SetSegmentType**( **_Index_** , **_SegmentType_** )

 _expression_ A variable that represents a **ShapeNodes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The node whose segment type is to be set.|
| _SegmentType_|Required| **[MsoSegmentType](http://msdn.microsoft.com/library/1a015227-8090-52a7-24f9-71d7e34fd05d%28Office.15%29.aspx)**|Specifies if the segment is straight or curved.|

## Example

This example changes all straight segments to curved segments in shape three on  `myDocument`. Shape three must be a freeform drawing.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Nodes 
    n = 1 
    While n <= .Count 
        If .Item(n).SegmentType = msoSegmentLine Then 
            .SetSegmentType n, msoSegmentCurve 
        End If 
        n = n + 1 
    Wend 
End With
```


## See also


#### Concepts


[ShapeNodes Object](shapenodes-object-excel.md)

