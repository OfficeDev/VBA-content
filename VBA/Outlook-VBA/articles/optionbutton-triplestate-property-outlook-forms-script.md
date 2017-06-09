---
title: OptionButton.TripleState Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 7643b4e7-1743-befd-9894-bee351296b79
ms.date: 06/08/2017
---


# OptionButton.TripleState Property (Outlook Forms Script)

Returns or sets a  **Boolean** that determines whether the **[OptionButton](optionbutton-object-outlook-forms-script.md)** supports the **Null** state. Read/write.


## Syntax

 _expression_. **TripleState**

 _expression_A variable that represents an  **OptionButton** object.


## Remarks

 **True** if the control clicks through three states, **False** if the control only supports two states, **True** and **False** (default).

Although the  **TripleState** property exists on the **OptionButton**, the property does not affect the action of the control. Regardless of the value of  **TripleState**, you cannot set the control to  **Null** through the user interface.

Regardless of the property setting, the  **Null** value can always be assigned programmatically to an **OptionButton**, causing that control to appear shaded.

When the  **TripleState** property is **True**, a user can choose from the values of  **Null**,  **True**, and  **False**. The null value is displayed as a shaded button.

When  **TripleState** is **False**, the user can choose either  **True** or **False**.

A control set to  **Null** does not initiate the **[Click](optionbutton-click-event-outlook-forms-script.md)** event.


