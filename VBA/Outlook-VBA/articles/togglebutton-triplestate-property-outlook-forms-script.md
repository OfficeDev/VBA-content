---
title: ToggleButton.TripleState Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: a82cbddf-3273-db90-57f7-26d12dac0c23
ms.date: 06/08/2017
---


# ToggleButton.TripleState Property (Outlook Forms Script)

Returns or sets a  **Boolean** that determines whether a user can specify, from the user interface, the **Null** state for a **[ToggleButton](togglebutton-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **TripleState**

 _expression_A variable that represents a  **ToggleButton** object.


## Remarks

 **True** if the control clicks through three states, **False** if the control only supports two states, **True** and **False** (default).

When the  **TripleState** property is **True**, a user can choose from the values of  **Null**,  **True**, and  **False**. The null value is displayed as a shaded button.

When  **TripleState** is **False**, the user can choose either  **True** or **False**.

A control set to  **Null** does not initiate the **[Click](togglebutton-click-event-outlook-forms-script.md)** event.

Regardless of the property setting, the  **Null** value can always be assigned programmatically to a **ToggleButton**, causing that control to appear shaded.


