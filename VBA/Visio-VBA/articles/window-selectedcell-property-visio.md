---
title: Window.SelectedCell Property (Visio)
keywords: vis_sdr.chm11660125
f1_keywords:
- vis_sdr.chm11660125
ms.prod: visio
api_name:
- Visio.Window.SelectedCell
ms.assetid: 104a2b2d-eb12-2917-6332-9a60e4623e74
ms.date: 06/08/2017
---


# Window.SelectedCell Property (Visio)

Returns the selected cell in the ShapeSheet window. Read-only.


## Syntax

 _expression_ . **SelectedCell**

 _expression_ A variable that represents a **Window** object.


### Return Value

Cell


## Remarks

The  **SelectedCell** property applies only to ShapeSheet windows. If you try to access the **SelectedCell** property for any other type of window, Microsoft Visio returns the error message "Invalid window type for this action."

If a ShapeSheet row is selected (instead of a cell),  **SelectedCell** returns **Nothing** . See the following example.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SelectedCell** property to print the name, section, row, column, and formula of the selected ShapeSheet cell in the Immediate window.


```vb
Public Sub SelectedCell_Example() 
 
 Dim vsoCell As Visio.Cell 
 
 Set vsoCell = Application.ActiveWindow.SelectedCell 
 
 'If vsoCell is Nothing, a row is probably selected. 
 If (Not vsoCell Is Nothing) Then 
 Debug.Print "Cell Name: " &; vsoCell.Name 
 Debug.Print "Section: " &; vsoCell.Section 
 Debug.Print "Row: " &; vsoCell.Row 
 Debug.Print "Column: " &; vsoCell.Column 
 Debug.Print "Formula: " &; vsoCell.Formula 
 Else 
 Debug.Print "vsoCell is Nothing--a row is probably selected." 
 
 End If 
 
End Sub
```


