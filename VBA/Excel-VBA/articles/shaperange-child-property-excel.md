---
title: ShapeRange.Child Property (Excel)
keywords: vbaxl10.chm640130
f1_keywords:
- vbaxl10.chm640130
ms.prod: excel
api_name:
- Excel.ShapeRange.Child
ms.assetid: ce25e66e-6446-1c43-1ab5-0ec486311ef2
ms.date: 06/08/2017
---


# ShapeRange.Child Property (Excel)

Returns  **msoTrue** if the specified shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **Child**

 _expression_ A variable that represents a **ShapeRange** object.


## Remarks





| **MsoTriState** can be one of these **MsoTriState** constants.|
| **msoCTrue** . Does not apply to this property.|
| **msoFalse** . If the selected shape is not a child shape.|
| **msoTriStateMixed** . If only some of the selected shapes are child shapes.|
| **msoTriStateToggle** . Does not apply to this property.|
| **msoTrue** . If the selected shape is a child shape.|

## Example

This example selects the first shape in the canvas, and if the selected shape is a child shape, fills the shape with the specified color. This example assumes that a drawing canvas contains multiple shapes on the active worksheet.


```vb
Sub FillChildShape() 
 
    'Select the first shape in the drawing canvas. 
    ActiveSheet.Shapes(1).CanvasItems(1).Select 
 
    'Fill selected shape if it is a child shape. 
    If Selection.ShapeRange.Child = msoTrue Then 
        Selection.ShapeRange.Fill.ForeColor.RGB = RGB(100, 0, 200) 
    Else 
        MsgBox "This shape is not a child shape." 
    End If 
 
End Sub
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

