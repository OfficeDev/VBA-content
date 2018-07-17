---
title: MultiPage.MultiRow Property (Outlook Forms Script)
keywords: olfm10.chm2001570
f1_keywords:
- olfm10.chm2001570
ms.prod: outlook
ms.assetid: 80375220-7268-f3a9-297e-29999fd3b3e3
ms.date: 06/08/2017
---


# MultiPage.MultiRow Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether the control has more than one row of tabs. Read/write.


## Syntax

 _expression_. **MultiRow**

 _expression_A variable that represents a  **MultiPage** object.


## Remarks

 **True** to allow more than one row of tabs, **False** to restrict tabs to a single row (default).

The width and number of tabs determines the number of rows. Changing the control's size also changes the number of rows. This allows the developer to resize the control and ensure that tabs wrap to fit the control. If the  **MultiRow** property is **False**, then truncation occurs if the width of the tabs exceeds the width of the control.


