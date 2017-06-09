---
title: Form.SelectionChange Event (Access)
keywords: vbaac10.chm13672
f1_keywords:
- vbaac10.chm13672
ms.prod: access
api_name:
- Access.Form.SelectionChange
ms.assetid: 4c815a6d-4971-6cbd-16ad-905e93ec1b52
ms.date: 06/08/2017
---


# Form.SelectionChange Event (Access)

Occurs whenever the user makes a new selection in a PivotChart view or PivotTable view.


## Syntax

 _expression_. **SelectionChange**

 _expression_ A variable that represents a **Form** object.


### Return Value

nothing


## Remarks

The user cannot cancel this event.


## Example

The following example demonstrates the syntax for a subroutine that traps the  **SelectionChange** event.


```vb
Private Sub Form_SelectionChange() 
 MsgBox "The selection has changed!" 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

