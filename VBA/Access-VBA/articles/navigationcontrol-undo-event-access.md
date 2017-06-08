---
title: NavigationControl.Undo Event (Access)
keywords: vbaac10.chm14210
f1_keywords:
- vbaac10.chm14210
ms.prod: access
api_name:
- Access.NavigationControl.Undo
ms.assetid: ebab443e-6abc-ed4a-5f2a-4ad00c7f9d8c
ms.date: 06/08/2017
---


# NavigationControl.Undo Event (Access)

Occurs when the user undoes a change.


## Syntax

 _expression_. **Undo**( ** _Cancel_**, )

 _expression_ A variable that represents a **NavigationControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**||

### Return Value

nothing


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


[NavigationControl Object](navigationcontrol-object-access.md)

