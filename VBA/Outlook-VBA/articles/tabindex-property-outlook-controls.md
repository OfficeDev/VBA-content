---
title: TabIndex Property (Outlook Controls)
keywords: olfm10.chm2002010
f1_keywords:
- olfm10.chm2002010
ms.prod: outlook
ms.assetid: cef32d27-35a6-28b5-657f-0ea1bcb8e10d
ms.date: 06/08/2017
---


# TabIndex Property (Outlook Controls)

Returns or sets an  **Integer** that specifies the position of a control in the form's tab order. Read/write.


## Syntax

 _expression_. **TabIndex**

 _expression_A variable that represents an Outlook control object.


## Remarks

The  **TabIndex** is an integer from 0 to one less than the number of controls on the form that have a **TabIndex** property. Assigning a **TabIndex** value of less than 0 generates an error. If you assign a **TabIndex** value greater than the largest index value, the system resets the value to the maximum allowable value.

The index value of the first object in the tab order is zero.


