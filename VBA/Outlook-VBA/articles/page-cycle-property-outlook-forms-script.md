---
title: Page.Cycle Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 729e72fa-5d2b-a754-481b-a9557d5d3c87
ms.date: 06/08/2017
---


# Page.Cycle Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies whether cycling includes controls nested in a **[MultiPage](multipage-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **Cycle**

 _expression_A variable that represents a  **Page** object.


## Remarks

The possible values for  **Cycle** are 0 and 2. 0 represents cycling through the controls on the form and the controls of the **MultiPage** that are currently displayed on the form. 2 represents cycling through the controls on the form and the **MultiPage**. The focus stays within the form and the  **MultiPage** until the focus is explicitly set to a control outside the form and the **MultiPage**.

If you specify a non-integer value for  **Cycle**, the value is rounded up to the nearest integer.

The tab order identifies the order in which controls receive the focus as the user tabs through a form or subform. The  **Cycle** property determines the action to take when a user tabs from the last control in the tab order.

The 0 setting transfers the focus to the first control of the next  **MultiPage** on the form when the user tabs from the last control in the tab order.

The 2 setting transfers the focus to the first control of the same form or the  **MultiPage** when the user tabs from the last control in the tab order.


