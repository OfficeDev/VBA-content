---
title: Form.ViewChange Event (Access)
keywords: vbaac10.chm13684
f1_keywords:
- vbaac10.chm13684
ms.prod: access
api_name:
- Access.Form.ViewChange
ms.assetid: a3788eca-783f-cb5d-1a7b-1c4a23648629
ms.date: 06/08/2017
---


# Form.ViewChange Event (Access)

Occurs whenever the specified PivotChart view or PivotTable view is redrawn.


## Syntax

 _expression_. **ViewChange**( ** _Reason_** )

 _expression_ A variable that represents a **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Reason_|Required|**Long**| A **PivotViewReasonEnum** constant that indicates how the view was changed. _Reason_ always returns ?1 for PivotChart Views.|

## Example

The following example demonstrates the syntax for a subroutine that traps the  **ViewChange** event.


```vb
Private Sub Form_ViewChange(ByVal Reason As Long) 
 If Reason = OWC.plViewReasonShowDetails Then 
 MsgBox "You've opted to show details." 
 End If 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

