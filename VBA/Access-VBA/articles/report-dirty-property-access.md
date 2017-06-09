---
title: Report.Dirty Property (Access)
keywords: vbaac10.chm13770
f1_keywords:
- vbaac10.chm13770
ms.prod: access
api_name:
- Access.Report.Dirty
ms.assetid: d9d9fe57-4fc5-9700-fc14-71f9eaa4a01b
ms.date: 06/08/2017
---


# Report.Dirty Property (Access)

You can use the  **Dirty** property to determine whether the current record has been modified since it was last saved. Read/write **Boolean**.


## Syntax

 _expression_. **Dirty**

 _expression_ A variable that represents a **Report** object.


## Remarks

For example, you may want to ask the user whether changes to a record were intended and, if not, allow the user to move to the next record without saving the changes. 

When a record is saved, Microsoft Access sets the  **Dirty** property to **False**. When a user makes changes to a record, the property is set to **True**.


## Example

The following example enables the  `btnUndo` button when data is changed. The UndoEdits( ) subroutine is called from the AfterUpdate event of text box controls. Clicking the enabled `btnUndo` button restores the original value of the control by using the **OldValue** property.


```vb
Sub UndoEdits() 
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


[Report Object](report-object-access.md)

