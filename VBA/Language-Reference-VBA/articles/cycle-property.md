---
title: Cycle Property
keywords: fm20.chm5225029
f1_keywords:
- fm20.chm5225029
ms.prod: office
api_name:
- Office.Cycle
ms.assetid: 8521df23-57d6-bcec-6d4e-ff77991b26f4
ms.date: 06/08/2017
---


# Cycle Property



Specifies the action to take when the user leaves the last control on a  **Frame** or **Page**.
 **Syntax**
 _object_. **Cycle** [= _fmCycle_ ]
The  **Cycle** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmCycle_|Optional. Specifies whether cycling includes controls nested in a  **Frame** or **MultiPage**.|
 **Settings**
The settings for  _fmCycle_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmCycleAllForms_|0|[Cycles](glossary-vba.md) through the controls on the form and the controls of the **Frame** and **MultiPage** controls that are currently displayed on the form.|
| _fmCycleCurrentForm_|2|Cycles through the controls on the form,  **Frame**, or **MultiPage**. The focus stays within the form, **Frame**, or **MultiPage** until the focus is explicitly set to a control outside the form, **Frame**, or **MultiPage**.|
If you specify a non-integer value for  **Cycle**, the value is rounded up to the nearest integer.
 **Remarks**
The [tab order](vbe-glossary.md) identifies the order in which controls receive the[focus](vbe-glossary.md) as the user tabs through a form or subform. The **Cycle** property determines the action to take when a user tabs from the last control in the tab order.
The  **fmCycleAllForms** setting transfers the focus to the the first control of the next **Frame** or **MultiPage** on the form when the user tabs from the last control in the tab order.
The  **fmCycleCurrentForm** setting transfers the focus to the the first control of the same form, **Frame**, or **MultiPage** when the user tabs from the last control in the tab order.

