---
title: Frame.Cycle Property (Outlook Forms Script)
keywords: olfm10.chm2001060
f1_keywords:
- olfm10.chm2001060
ms.prod: outlook
ms.assetid: 012c4b16-8c4d-fd11-39cc-9fe1799630c8
ms.date: 06/08/2017
---


# Frame.Cycle Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies whether cycling includes controls nested in a **[Frame](frame-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **Cycle**

 _expression_A variable that represents a  **Frame** object.


## Remarks

The possible values for  **Cycle** are 0 and 2. 0 represents cycling through the controls on the form and the controls of the **Frame** that are currently displayed on the form. 2 represents cycling through the controls on the form and the **Frame**. The focus stays within the form and the  **Frame** until the focus is explicitly set to a control outside the form and the **Frame**.

If you specify a non-integer value for  **Cycle**, the value is rounded up to the nearest integer.

The tab order identifies the order in which controls receive the focus as the user tabs through a form or subform. The  **Cycle** property determines the action to take when a user tabs from the last control in the tab order.

The 0 setting transfers the focus to the first control of the next  **Frame** on the form when the user tabs from the last control in the tab order.

The 2 setting transfers the focus to the first control of the same form or the  **Frame** when the user tabs from the last control in the tab order.


