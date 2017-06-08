---
title: Form.Dirty Event (Access)
keywords: vbaac10.chm13662
f1_keywords:
- vbaac10.chm13662
ms.prod: access
api_name:
- Access.Form.Dirty
ms.assetid: e0bcf968-7176-bd24-29c4-d3f014f57adb
ms.date: 06/08/2017
---


# Form.Dirty Event (Access)

The Dirty event occurs when the contents of the specified control changes.


## Syntax

 _expression_. **Dirty**( ** _Cancel_**, )

 _expression_ A variable that represents a **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**|The setting determines if the  **Dirty** event occurs. Setting the _Cancel_ argument to **True** (?1) cancels the **Dirty** event.|

## Remarks

Examples of this event include entering a character directly in the text box or combo box or changing the control's  **Text** property setting by using a macro or Visual Basic.


- Modifying a record within a form by using a macro or Visual Basic doesn't trigger this event. You must type the data directly into the record or set the control's  **Text** property.
    
- This event applies only to bound forms, not an unbound form or report.
    
To run a macro or event procedure when this event occurs, set the  **OnDirty** property to the name of the macro or to [Event Procedure].

By running a macro or event procedure when a Dirty event occurs, you can determine if the record can be changed. You can also display a message and ask for edit permission.

Changing the data in a record by using the keyboard causes keyboard events to occur in addition to control events like the Dirty event. For example, if you move to a new record and type an ANSI character in a text box in the record, the following events occur in this order:

 **KeyDown** > **KeyPress** > **BeforeInsert** > **Dirty** > **KeyUp**

The  **BeforeUpdate** and **AfterUpdate** events for a record occur after you have entered the new or changed data in the record and moved to another record (or clicked **Save Record** on the **Records** menu), and therefore after the Dirty event for the record.

Canceling the Dirty event will cause the changes to the current record to be rolled back. It is equivalent to pressing the ESC key.


## Example

The following example enables the btnUndo button when data is changed. The UndoEdits( ) subroutine is called from the Dirty event of text box controls. Clicking the enabled btnUndo button restores the original value of the control by using the  **OldValue** property.


```vb
Private Sub Form_Dirty() 
 If Me.Dirty Then 
 Me!btnUndo.Enabled = True ' Enable button. 
 Else 
 Me!btnUndo.Enabled = False ' Disable button. 
 End If 
End Sub 
 
Sub btnUndo_Click() 
 Dim ctlC As Control 
 ' For each control. 
 For Each ctlC in Me.Controls 
 If ctlC.ControlType = acTextBox Then 
 ' Restore Old Value. 
 ctlC.Value = ctlC.OldValue 
 End If 
 Next ctlC 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

