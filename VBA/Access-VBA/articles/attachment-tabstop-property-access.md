---
title: Attachment.TabStop Property (Access)
keywords: vbaac10.chm14013
f1_keywords:
- vbaac10.chm14013
ms.prod: access
api_name:
- Access.Attachment.TabStop
ms.assetid: 918d81a6-a9a2-ab4e-6fb3-ad78233b6e7f
ms.date: 06/08/2017
---


# Attachment.TabStop Property (Access)

You can use the  **TabStop** property to specify whether you can use the TAB key to move the focus to a control in Form view. Read/write **Boolean**.


## Syntax

 _expression_. **TabStop**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

The  **TabStop** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|(Default) You can move the focus to the control by pressing the TAB key.|
|No|**False**|You cannot move the focus to the control by pressing the TAB key.|
When you create a control on a form, Microsoft Access automatically assigns the control a position in the form's tab order. Each new control is placed last in the tab order. If you want to prevent a control from being available when you tab through the controls in a form, set the control's  **TabStop** property to No.

In Form view, hidden or disabled controls remain in the tab order but are skipped when you move through the controls by pressing TAB, even if their  **TabStop** properties are set to Yes.

As long as a control's  **Enabled** property is set to Yes, you can click the control or use an access key to select it, regardless of its **TabStop** property setting. For example, you can set the **TabStop** property of a command button to No to prevent users from selecting the button by pressing TAB. However, they can still click the command button to select it.


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

