---
title: Selection.Align Method (Visio)
keywords: vis_sdr.chm11151435
f1_keywords:
- vis_sdr.chm11151435
ms.prod: visio
api_name:
- Visio.Selection.Align
ms.assetid: 4a73dfee-2a78-f459-4481-5f722feb7204
ms.date: 06/08/2017
---


# Selection.Align Method (Visio)

Aligns two or more selected shapes.


## Syntax

 _expression_ . **Align**( **_AlignHorizontal_** , **_AlignVertical_** , **_GlueToGuide_** )

 _expression_ A variable that represents a **Selection** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _AlignHorizontal_|Required| **VisHorizontalAlignTypes**|Aligns selected shapes along a horizontal axis. See Remarks for possible values.|
| _AlignVertical_|Required| **VisVerticalAlignTypes**|Aligns selected shapes along a vertical axis. See Remarks for possible values.|
| _GlueToGuide_|Optional| **Boolean**|If  **True** , creates a guide and glues selected shapes to it; if **False** , it does not. The default is **False** .|

### Return Value

Nothing


## Remarks

The following possible values for  _AlignHorizontal_ are declared in **VisHorizontalSelectionTypes** in the Visio type library.



|** Constant**|** Value**|** Description**|
|:-----|:-----|:-----|
| **visHorzAlignCenter**|2|Aligns to the center of the primary selected shape.|
| **visHorzAlignLeft**|1|Aligns to the left of the primary selected shape.|
| **visHorzAlignNone**|0|Does not align horizontally.|
| **visHorzAlignRight**|3|Aligns to the right of the primary selected shape.|
The following possible values for  _AlignVertical_ are declared in **VisVerticalSelectionTypes** in the Visio type library.



|** Constant**|** Value**|** Description**|
|:-----|:-----|:-----|
| **visVertAlignBottom**|3|Aligns to bottom of primary selected shape.|
| **visVertAlignMiddle**|2|Aligns to middle of primary selected shape.|
| **visVertAlignNone**|0|Does not align vertically. |
| **visVertAlignTop**|1|Aligns to top of primary selected shape.|
If you pass non-zero values for both  _AlignHorizontal_ and _AlignVertical_, the selected shapes appear superimposed. The most recently created shape appears at the front of the z-order.

Calling the  **Align** method is equivalent to clicking **Position** on the **Home** tab and then setting options under **Align Shapes**. 


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Align** method to align three shapes vertically.


```vb
Public Sub Align_Example() 
 
    Dim vsoShape1 As Visio.Shape 
    Dim vsoShape2 As Visio.Shape 
    Dim vsoShape3 As Visio.Shape 
     
    Set vsoShape1 = Application.ActiveWindow.Page.DrawRectangle(1, 9, 3, 7) 
    Set vsoShape2 = Application.ActiveWindow.Page.DrawRectangle(3, 6, 5, 5) 
    Set vsoShape3 = Application.ActiveWindow.Page.DrawRectangle(6, 4, 8, 2) 
 
    ActiveWindow.DeselectAll 
     
    ActiveWindow.Select vsoShape1, visSelect 
    ActiveWindow.Select vsoShape2, visSelect 
    ActiveWindow.Select vsoShape3, visSelect 
     
    Application.ActiveWindow.Selection.Align visHorzAlignRight, visVertAlignNone, False 
 
End Sub
```


