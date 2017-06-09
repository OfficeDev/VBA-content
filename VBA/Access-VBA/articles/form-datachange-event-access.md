---
title: Form.DataChange Event (Access)
keywords: vbaac10.chm13685
f1_keywords:
- vbaac10.chm13685
ms.prod: access
api_name:
- Access.Form.DataChange
ms.assetid: 026fddb4-2a43-095c-9460-98c12378735c
ms.date: 06/08/2017
---


# Form.DataChange Event (Access)

Occurs when certain properties are changed or when certain methods are executed in the specified PivotTable view.


## Syntax

 _expression_. **DataChange**( ** _Reason_**, )

 _expression_ A variable that represents a **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Reason_|Required|**Long**|A  **PivotDataReasonEnum** constant that indicates the reason that this event was triggered.|

### Return Value

nothing


## Example

The following example demonstrates the syntax for a subroutine that traps the  **DataChange** event.






```vb
Private Sub Form_DataChange(Reason As Long) 
 If Reason = OWC.plDataReasonDisplayCellColorChange Then 
 MsgBox "The cell display color was changed." 
 End If 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

