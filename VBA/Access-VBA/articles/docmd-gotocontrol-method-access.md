---
title: DoCmd.GoToControl Method (Access)
keywords: vbaac10.chm4152
f1_keywords:
- vbaac10.chm4152
ms.prod: access
api_name:
- Access.DoCmd.GoToControl
ms.assetid: 2b51231d-f6a4-4891-d49d-bedb68f85b04
ms.date: 06/08/2017
---


# DoCmd.GoToControl Method (Access)

The  **GoToControl** method performs the GoToControl action action in Visual Basic.


## Syntax

 _expression_. **GoToControl**( ** _ControlName_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ControlName_|Required|**Variant**|A string expression that is the name of a control on the active form or datasheet.|

## Remarks

You can use the GoToControl method to move the focus to the specified field or control in the current record of the open form, form datasheet, table datasheet, or query datasheet. You can use this method when you want a particular field or control to have the focus. This field or control can then be used for comparisons or  **FindRecord** actions. You can also use this method to navigate in a form according to certain conditions. For example, if the user enters No in a Married control on a health insurance form, the focus can automatically skip the Spouse/partner Name control and move to the next control.

You cannot use the  **GoToControl** method to move the focus to a control on a hidden form.

Use only the name of the control for the  _controlname_ argument, not the full syntax.

You can also use the  **SetFocus** method to move the focus to a control on a form or any of its subforms, or to a field in an open table, query, or form datasheet. This is the preferred method for moving the focus in Visual Basic, especially to controls on subforms and nested subforms, because you can use the full syntax to specify the control you want to move to.

You can use the  **GoToControl** method to move to a subform, which is a type of control. You can then use the **GoToControl** method to move to a particular record in the subform. You can also move to a control on a subform by using the **GoToControl** method to move first to the subform and then to the control on the subform.


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

