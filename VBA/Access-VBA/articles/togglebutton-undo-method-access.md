---
title: ToggleButton.Undo Method (Access)
keywords: vbaac10.chm11689
f1_keywords:
- vbaac10.chm11689
ms.prod: access
api_name:
- Access.ToggleButton.Undo
ms.assetid: 95db5f79-af3d-9577-8d7e-6d2784a016f4
ms.date: 06/08/2017
---


# ToggleButton.Undo Method (Access)

You can use the  **Undo** method to reset a control or form when its value has been changed.


## Syntax

 _expression_. **Undo**

 _expression_ A variable that represents a **ToggleButton** object.


## Remarks

For example, you can use the  **Undo** method to clear a change to a record that contains an invalid entry.

If the  **Undo** method is applied to a form, all changes to the current record are lost. If the **Undo** method is applied to a control, only the control itself is affected.

This method must be applied before the form or control is updated. You may want to include this method in a form's  **BeforeUpdate** event or in a control's **Change** event.

The  **Undo** method offers an alternative to using the **SendKeys** statement to send the value of the ESC key in an event procedure.


## Example

The following example shows how you can use the  **Undo** method within a control's **Change** event procedure to force a field named LastName to reset to its original value, if it changed.


```vb
Private Sub LastName_Change() 
 Me!LastName.Undo 
End Sub
```

The next example uses the  **Undo** method to reset all changes to a form before the form is updated.




```vb
Private Sub Form_BeforeUpdate(Cancel As Integer) 
 Me.Undo 
End Sub
```


## See also


#### Concepts


[ToggleButton Object](togglebutton-object-access.md)

