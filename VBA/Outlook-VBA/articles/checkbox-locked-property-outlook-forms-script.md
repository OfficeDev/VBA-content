---
title: CheckBox.Locked Property (Outlook Forms Script)
keywords: olfm10.chm2001470
f1_keywords:
- olfm10.chm2001470
ms.prod: outlook
ms.assetid: 7cc183ed-44d1-8407-919a-43de05d52e75
ms.date: 06/08/2017
---


# CheckBox.Locked Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether a control can be edited. Read/write.


## Syntax

 _expression_. **Locked**

 _expression_A variable that represents a  **CheckBox** object.


## Remarks

 **True** represents that the value of the control cannot be edited. **False** represents that the value can be edited (default).

When a control is locked and enabled, it can still initiate events and can still receive the focus.

The  **[Enabled](checkbox-enabled-property-outlook-forms-script.md)** and **Locked** properties work together to achieve the following effects:


- If  **Enabled** and **Locked** are both **True**, the control can receive focus and appears normally (not dimmed) in the form. The user can copy, but not edit, data in the control.
    
- If  **Enabled** is **True** and **Locked** is **False**, the control can receive focus and appears normally in the form. The user can copy and edit data in the control.
    
- If  **Enabled** is **False** and **Locked** is **True**, the control cannot receive focus and is dimmed in the form. The user can neither copy nor edit data in the control.
    
- If  **Enabled** and **Locked** are both **False**, the control cannot receive focus and is dimmed in the form. The user can neither copy nor edit data in the control.
    



