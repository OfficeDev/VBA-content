---
title: TextBox.Undo Event (Access)
keywords: vbaac10.chm14210
f1_keywords:
- vbaac10.chm14210
ms.prod: access
api_name:
- Access.TextBox.Undo
ms.assetid: ee009e53-41be-0c9a-a92d-15572f6213b6
ms.date: 06/08/2017
---


# TextBox.Undo Event (Access)

Occurs when the user undoes a change.


## Syntax

 _expression_. **Undo**( ** _Cancel_**, )

 _expression_ A variable that represents a **TextBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**|Set this argument to  **True** to cancel the undo operation and leave the control or form in its edited state.|

## Remarks

The Undo event for controls occurs whenever the user returns a control to its original state by clicking the  **Undo Field/Record** button on the command bar, clicking the **Undo** button, pressing the ESC key, or calling the **Undo** method of the specified control. The control needs to have focus in all three cases. The event does not occur if the user clicks the **Undo Typing** button on the command bar.


## Example

The following example demonstrates the syntax for a subroutine that traps the Undo event for a form.


```vb
Private Sub Form_Undo(Cancel As Integer) 
 Dim intResponse As Integer 
 Dim strPrompt As String 
 
 strPrompt = "Cancel the undo operation?" 
 
 intResponse = MsgBox(strPrompt, vbYesNo) 
 
 If intResponse = vbYes Then 
 Cancel = True 
 Else 
 Cancel = False 
 End If 
End Sub
```


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

