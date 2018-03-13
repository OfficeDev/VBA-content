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


| <strong>Part</strong> | <strong>Description</strong>                                                                                             |
|:----------------------|:-------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. A valid object.                                                                                                |
| <em>fmCycle</em>      | Optional. Specifies whether cycling includes controls nested in a  <strong>Frame</strong> or <strong>MultiPage</strong>. |

 **Settings**
The settings for  _fmCycle_ are:


| <strong>Constant</strong>   | <strong>Value</strong> | <strong>Description</strong>                                                                                                                                                                                                                                                                                    |
|:----------------------------|:-----------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>fmCycleAllForms</em>    | 0                      | [Cycles](glossary-vba.md) through the controls on the form and the controls of the <strong>Frame</strong> and <strong>MultiPage</strong> controls that are currently displayed on the form.                                                                                                                     |
| <em>fmCycleCurrentForm</em> | 2                      | Cycles through the controls on the form,  <strong>Frame</strong>, or <strong>MultiPage</strong>. The focus stays within the form, <strong>Frame</strong>, or <strong>MultiPage</strong> until the focus is explicitly set to a control outside the form, <strong>Frame</strong>, or <strong>MultiPage</strong>. |

If you specify a non-integer value for  **Cycle**, the value is rounded up to the nearest integer.
 **Remarks**
The [tab order](vbe-glossary.md) identifies the order in which controls receive the[focus](vbe-glossary.md) as the user tabs through a form or subform. The **Cycle** property determines the action to take when a user tabs from the last control in the tab order.
The  **fmCycleAllForms** setting transfers the focus to the the first control of the next **Frame** or **MultiPage** on the form when the user tabs from the last control in the tab order.
The  **fmCycleCurrentForm** setting transfers the focus to the the first control of the same form, **Frame**, or **MultiPage** when the user tabs from the last control in the tab order.

