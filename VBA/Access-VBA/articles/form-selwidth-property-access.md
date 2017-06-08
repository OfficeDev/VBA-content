---
title: Form.SelWidth Property (Access)
keywords: vbaac10.chm13471
f1_keywords:
- vbaac10.chm13471
ms.prod: access
api_name:
- Access.Form.SelWidth
ms.assetid: a5ce22e3-af69-209c-f988-16cf4f77fd62
ms.date: 06/08/2017
---


# Form.SelWidth Property (Access)

You can use the  **SelWidth** property to specify or determine the number of selected columns (fields) in the current selection rectangle. Read/write **Long**.


## Syntax

 _expression_. **SelWidth**

 _expression_ A variable that represents a **Form** object.


## Remarks

If there's no selection, the value returned by this property will be zero. Setting this property to 0 removes the selection from the datasheet or form.

If you've selected one or more records in the datasheet (using the record selectors), you can't change the setting of the  **SelWidth** property (except to set it to 0).

You can use these properties with the  **SelTop** and **SelLeft** properties to specify or determine the actual position of the selection rectangle on the datasheet. If there's no selection, then the **SelTop** and **SelLeft** properties return the row number and column number of the cell with the focus.

The  **SelHeight** and **SelWidth** properties contain the position of the lower-right corner of the selection rectangle. The **SelTop** and **SelLeft** property values determine the upper-left corner of the selection rectangle.


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

