---
title: TabStrip.MultiRow Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 09dc5bcc-4425-8f37-24fa-3b74af0e4605
ms.date: 06/08/2017
---


# TabStrip.MultiRow Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether the control has more than one row of tabs. Read/write.


## Syntax

 _expression_. **MultiRow**

 _expression_A variable that represents a  **TabStrip** object.


## Remarks

 **True** to allow more than one row of tabs, **False** to restrict tabs to a single row (default).

The width and number of tabs determines the number of rows. Changing the control's size also changes the number of rows. This allows the developer to resize the control and ensure that tabs wrap to fit the control. If the  **MultiRow** property is **False**, then truncation occurs if the width of the tabs exceeds the width of the control.

If  **MultiRow** is **False** and tabs are truncated, there will be a small scroll bar on the **[TabStrip](tabstrip-object-outlook-forms-script.md)** to allow scrolling to the other tabs or pages.


