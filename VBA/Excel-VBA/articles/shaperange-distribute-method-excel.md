---
title: ShapeRange.Distribute Method (Excel)
keywords: vbaxl10.chm640080
f1_keywords:
- vbaxl10.chm640080
ms.prod: excel
api_name:
- Excel.ShapeRange.Distribute
ms.assetid: cef14a4b-4d6e-758e-928a-99233f893ddc
ms.date: 06/08/2017
---


# ShapeRange.Distribute Method (Excel)

Horizontally or vertically distributes the shapes in the specified range of shapes.


## Syntax

 _expression_ . **Distribute**( **_DistributeCmd_** , **_RelativeTo_** )

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DistributeCmd_|Required| **[MsoDistributeCmd](http://msdn.microsoft.com/library/be0cb580-1d82-1443-0f03-3461ddb9c33d%28Office.15%29.aspx)**|Specifies whether shapes in the range are to be distributed horizontally or vertically.|
| _RelativeTo_|Required| **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)**|Not used in Microsoft Excel. Must be  **False** .|

## Example

This example defines a shape range that contains all the AutoShapes on  `myDocument` and then horizontally distributes the shapes in this range. The leftmost shape retains its position.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
    numShapes = .Count 
    If numShapes > 1 Then 
        numAutoShapes = 0 
        ReDim autoShpArray(1 To numShapes) 
        For i = 1 To numShapes 
            If .Item(i).Type = msoAutoShape Then 
                numAutoShapes = numAutoShapes + 1 
                autoShpArray(numAutoShapes) = .Item(i).Name 
            End If 
        Next 
        If numAutoShapes > 1 Then 
            ReDim Preserve autoShpArray(1 To numAutoShapes) 
            Set asRange = .Range(autoShpArray) 
            asRange.Distribute msoDistributeHorizontally, False 
        End If 
    End If 
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

