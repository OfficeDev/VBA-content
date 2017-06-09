---
title: OptionButton.TabStop Property (Access)
keywords: vbaac10.chm10579
f1_keywords:
- vbaac10.chm10579
ms.prod: access
api_name:
- Access.OptionButton.TabStop
ms.assetid: b7bd7921-2ba3-1445-1e89-ce8fa0c2ed4e
ms.date: 06/08/2017
---


# OptionButton.TabStop Property (Access)

You can use the  **TabStop** property to specify whether you can use the TAB key to move the focus to a control. Read/write **Boolean**.


## Syntax

 _expression_. **TabStop**

 _expression_ A variable that represents an **OptionButton** object.


## Remarks

This property doesn't apply to toggle button controls when they appear in an option group. It applies only to the option group itself.

The  **TabStop** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|(Default) You can move the focus to the control by pressing the TAB key.|
|No|**False**|You can't move the focus to the control by pressing the TAB key.|
When you create a control on a form, Microsoft Access automatically assigns the control a position in the form's tab order. Each new control is placed last in the tab order. If you want to prevent a control from being available when you tab through the controls in a form, set the control's  **TabStop** property to No.

In Form view, hidden or disabled controls remain in the tab order but are skipped when you move through the controls by pressing TAB, even if their  **TabStop** properties are set to Yes.

As long as a control's  **Enabled** property is set to Yes, you can click the control or use an access key to select it, regardless of its **TabStop** property setting. For example, you can set the **TabStop** property of a command button to No to prevent users from selecting the button by pressing TAB. However, they can still click the command button to choose it.


## Example

The following example disables the ability to move the focus to the "City" text box on the "Suppliers" form by using the TAB key.


```vb
Forms("Suppliers").Controls("City").TabStop = False
```


## See also


#### Concepts


[OptionButton Object](optionbutton-object-access.md)

