---
title: Form.SelLeft Property (Access)
keywords: vbaac10.chm13469
f1_keywords:
- vbaac10.chm13469
ms.prod: access
api_name:
- Access.Form.SelLeft
ms.assetid: ddc05c0a-3132-5380-33c9-96fa2f92571d
ms.date: 06/08/2017
---


# Form.SelLeft Property (Access)

You can use the  **SelLeft** property to specify or determine which column (field) is leftmost in the current selection rectangle. Read/write **Long**.


## Syntax

 _expression_. **SelLeft**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **SelLeft** property returns a value between 1 and the number of columns in the datasheet.

If there's no selection, the value returned by these properties is the row and column of the cell with the focus. If you've selected one or more records in the datasheet (using the record selectors), you can't change the setting of the  **SelLeft** property.

You can use these properties with the  **SelHeight** and **SelWidth** properties to specify or determine the actual size of the selection rectangle. The **SelTop** and **SelLeft** properties determine the position of the upper-left corner of the selection rectangle. The **SelHeight** and **SelWidth** properties determine the lower-right corner of the selection rectangle.


## Example

The following example shows how to use the  **SelHeight**, **SelWidth**, **SelTop**, and **SelLeft** properties to determine the position and size of a selection rectangle in datasheet view. The SetHeightWidth procedure assigns the height and width of the current selection rectangle to the variables `lngNumRows`,  `lngNumColumns`,  `lngTopRow`, and  `lngLeftColumn`, and displays those values in a message box.


```vb
Public Sub SetHeightWidth(ByRef frm As Form) 
 
 Dim lngNumRows As Long 
 Dim lngNumColumns As Long 
 Dim lngTopRow As Long 
 Dim lngLeftColumn As Long 
 Dim strMsg As String 
 
 ' Form is in Datasheet view. 
 If frm.CurrentView = 2 Then 
 
 ' Number of rows selected. 
 lngNumRows = frm.SelHeight 
 
 ' Number of columns selected. 
 lngNumColumns = frm.SelWidth 
 
 ' Topmost row selected. 
 lngTopRow = frm.SelTop 
 
 ' Leftmost column selected. 
 lngLeftColumn = frm.SelLeft 
 
 ' Display message. 
 strMsg = "Number of rows: " &; lngNumRows &; vbCrLf 
 strMsg = strMsg &; "Number of columns: " _ 
 &; lngNumColumns &; vbCrLf 
 strMsg = strMsg &; "Top row: " &; lngTopRow &; vbCrLf 
 strMsg = strMsg &; "Left column: " &; lngLeftColumn 
 MsgBox strMsg, vbInformation 
 End If 
 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

