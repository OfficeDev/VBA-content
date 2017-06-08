---
title: Page.ControlTipText Property (Outlook Forms Script)
keywords: olfm10.chm2000990
f1_keywords:
- olfm10.chm2000990
ms.prod: outlook
ms.assetid: 11412cc8-7e62-1382-de69-905d5d75d419
ms.date: 06/08/2017
---


# Page.ControlTipText Property (Outlook Forms Script)

Returns and sets a  **String** that specifies text that appears when the user briefly holds the mouse pointer over a control without clicking. Read/write.


## Syntax

 _expression_. **ControlTipText**

 _expression_A variable that represents a  **Page** object.


## Remarks

The  **ControlTipText** property lets you give users tips about a control in a running form. The property can be set during design time but only appears by the control during run time.

The default value of  **ControlTipText** is an empty string. When the value of **ControlTipText** is set to an empty string, no tip is available for that control.


