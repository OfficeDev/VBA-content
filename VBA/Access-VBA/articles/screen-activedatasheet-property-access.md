---
title: Screen.ActiveDatasheet Property (Access)
keywords: vbaac10.chm12487
f1_keywords:
- vbaac10.chm12487
ms.prod: access
api_name:
- Access.Screen.ActiveDatasheet
ms.assetid: cff189e7-9b8a-280f-e287-e4367f8ac134
ms.date: 06/08/2017
---


# Screen.ActiveDatasheet Property (Access)

You can use the  **ActiveDatasheet** property together with the **[Screen](screen-object-access.md)** object to identify or refer to the datasheet that has the focus. Read-only **Form** object.


## Syntax

 _expression_. **ActiveDatasheet**

 _expression_ A variable that represents a **Screen** object.


## Remarks

The  **ActiveDatasheet** property setting contains the datasheet object that has the focus at run time.

You can use this property to refer to an active datasheet together with one of its properties or methods. For example, the following code uses the  **ActiveDatasheet** property to reference the top row of the selection in the active datasheet.




```vb
TopRow = Screen.ActiveDatasheet.SelTop
```


## Example

The following example uses the  **ActiveDatasheet** property to identify the datasheet cell with the focus, or if more than one cell is selected, the location of the first row and column in the selection.


```vb
Public Sub GetSelection() 
 ' This procedure demonstrates how to get a pointer to the 
 ' current active datasheet. 
 
 Dim objDatasheet As Object 
 Dim lngFirstRow As Long 
 Dim lngFirstColumn As Long 
 Const conNoActiveDatasheet = 2484 
 
 On Error GoTo GetSelection_Err 
 
 Set objDatasheet = Screen.ActiveDatasheet 
 
 lngFirstRow = objDatasheet.SelTop 
 lngFirstColumn = objDatasheet.SelLeft 
 MsgBox "The first item in this selection is located at " &; _ 
 "Row " &; lngFirstRow &; ", Column " &; _ 
 lngFirstColumn, vbInformation 
 
GetSelection_Bye: 
 Exit Sub 
GetSelection_Err: 
 If Err = conNoActiveDatasheet Then 
 MsgBox "No data sheet is active.", vbExclamation 
 Resume GetSelection_Bye 
 End If 
End Sub
```


## See also


#### Concepts


[Screen Object](screen-object-access.md)

