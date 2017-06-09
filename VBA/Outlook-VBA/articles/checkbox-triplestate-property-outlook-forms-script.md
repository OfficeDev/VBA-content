---
title: CheckBox.TripleState Property (Outlook Forms Script)
keywords: olfm10.chm2002150
f1_keywords:
- olfm10.chm2002150
ms.prod: outlook
ms.assetid: 6d68324c-a551-b0d4-b89e-28e1045f0992
ms.date: 06/08/2017
---


# CheckBox.TripleState Property (Outlook Forms Script)

Returns or sets a  **Boolean** that determines whether a user can specify, from the user interface, the **Null** state for a **[CheckBox](checkbox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **TripleState**

 _expression_A variable that represents a  **CheckBox** object.


## Remarks

 **True** if the control clicks through three states, **False** if the control only supports two states, **True** and **False** (default).

When the  **TripleState** property is **True**, a user can choose from the values of  **Null**,  **True**, and  **False**. The  **Null** value is displayed as a shaded button.

When  **TripleState** is **False**, the user can choose either  **True** or **False**.

A control set to  **Null** does not initiate the **[Click](checkbox-click-event-outlook-forms-script.md)** event.

Regardless of the property setting, the  **Null** value can always be assigned programmatically to a **CheckBox**, causing that control to appear shaded.


