---
title: ShapeNodes.Insert Method (Excel)
keywords: vbaxl10.chm112008
f1_keywords:
- vbaxl10.chm112008
ms.prod: excel
api_name:
- Excel.ShapeNodes.Insert
ms.assetid: b4f7e695-2102-5cbd-2d6b-bc167407cc0f
ms.date: 06/08/2017
---


# ShapeNodes.Insert Method (Excel)

Inserts a node into a freeform shape.


## Syntax

 _expression_ . **Insert**( **_Index_** , **_SegmentType_** , **_EditingType_** , **_X1_** , **_Y1_** , **_X2_** , **_Y2_** , **_X3_** , **_Y3_** )

 _expression_ A variable that represents a **ShapeNodes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Integer**| **Long** . The number of the shape node after which to insert a new node.|
| _SegmentType_|Required| **[MsoSegmentType](http://msdn.microsoft.com/library/1a015227-8090-52a7-24f9-71d7e34fd05d%28Office.15%29.aspx)**|The segment type.|
| _EditingType_|Required| **[MsoEditingType](http://msdn.microsoft.com/library/5fe5c4f6-6467-c6a7-197c-ff700c384b92%28Office.15%29.aspx)**|The editing type.|
| _X1_|Required| **Single**|If the  _EditingType_ of the new segment is **msoEditingAuto** , this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the end point of the new segment. If the _EditingType_ of the new node is **msoEditingCorner** , this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the first control point for the new segment.|
| _Y1_|Required| **Single**|If the  _EditingType_ of the new segment is **msoEditingAuto** , this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the end point of the new segment. If the _EditingType_ of the new node is **msoEditingCorner** , this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the first control point for the new segment.|
| _X2_|Required| **Single**|If the  _EditingType_ of the new segment is **msoEditingCorner** , this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the second control point for the new segment. If the _EditingType_ of the new segment is **msoEditingAuto** , don't specify a value for this argument.|
| _Y2_|Required| **Single**|If the  _EditingType_ of the new segment is **msoEditingCorner** , this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the second control point for the new segment. If the _EditingType_ of the new segment is **msoEditingAuto** , don't specify a value for this argument.|
| _X3_|Required| **Single**|If the  _EditingType_ of the new segment is **msoEditingCorner** , this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the end point of the new segment. If the _EditingType_ of the new segment is **msoEditingAuto** , don't specify a value for this argument.|
| _Y3_|Required| **Single**|If the  _EditingType_ of the new segment is **msoEditingCorner** , this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the end point of the new segment. If the _EditingType_ of the new segment is **msoEditingAuto** , don't specify a value for this argument.|

## Example

This example selects the third shape in the active document, checks whether the shape is a Freeform object, and if it is, inserts a node. This example assumes three shapes exist on the active worksheet.


```vb
Sub InsertShapeNode() 
    ActiveSheet.Shapes(3).Select 
    With Selection.ShapeRange 
        If .Type = msoFreeform Then 
            .Nodes.Insert _ 
                Index:=3, SegmentType:=msoSegmentCurve, _ 
                EditingType:=msoEditingSymmetric, X1:=35, Y1:=100 
            .Fill.ForeColor.RGB = RGB(0, 0, 200) 
            .Fill.Visible = msoTrue 
        Else 
            MsgBox "This shape is not a Freeform object." 
        End If 
    End With 
End Sub
```


## See also


#### Concepts


[ShapeNodes Object](shapenodes-object-excel.md)

