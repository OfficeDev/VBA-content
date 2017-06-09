---
title: BoundObjectFrame.OldValue Property (Access)
keywords: vbaac10.chm10898
f1_keywords:
- vbaac10.chm10898
ms.prod: access
api_name:
- Access.BoundObjectFrame.OldValue
ms.assetid: 89d4855e-9c7e-7c3f-4063-f9f74d7245ca
ms.date: 06/08/2017
---


# BoundObjectFrame.OldValue Property (Access)

You can use the  **OldValue** property to determine the unedited value of a bound control. Read-only **Variant**.


## Syntax

 _expression_. **OldValue**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

The  **OldValue** property contains the unedited data from a bound control and is read-only in all views.

Microsoft Access uses the  **OldValue** property to store the original value of a bound control. When you edit a bound control on a form, your changes aren't saved until you move to another record. The **OldValue** property contains the unedited version of the underlying data.

You can provide your own undo capability by assigning the  **OldValue** property setting to a control. The following example shows how you can undo any changes to text box controls on a form:




```vb
Private Sub btnUndo_Click() 
 
 Dim ctlTextbox As Control 
 
 For Each ctlTextbox in Me.Controls 
 If ctlTextbox.ControlType = acTextBox Then 
 ctlTextbox.Value = ctl.OldValue 
 End If 
 Next ctlTextbox 
 
End Sub
```

If the control hasn't been edited, this code has no effect. When you move to another record, the record source is updated, so the current value and the  **OldValue** property will be the same.

The  **OldValue** property setting has the same data type as the field to which the control is bound.


## Example

The following example checks to determine if new data entered in a field is within 10 percent of the value of the original data. If the change is greater than 10 percent, the  **OldValue** property is used to restore the original value. This procedure could be called from the BeforeUpdate event of the control that contains data you want to validate.


```vb
Public Sub Validate_Field() 
 
 Dim curNewValue As Currency 
 Dim curOriginalValue As Currency 
 Dim curChange As Currency 
 Dim strMsg As String 
 
 curNewValue = Forms!Products!UnitPrice 
 curOriginalValue = Forms!Products!UnitPrice.OldValue 
 curChange = Abs(curNewValue - curOriginalValue) 
 
 If curChange > (curOriginalValue * .1) Then 
 strMsg = "Change is more than 10% of original unit price. " _ 
 &; "Restoring original unit price." 
 MsgBox strMsg, vbExclamation, "Invalid change." 
 Forms!Products!UnitPrice = curOriginalValue 
 End If 
 
End Sub
```


## See also


#### Concepts


[BoundObjectFrame Object](boundobjectframe-object-access.md)

