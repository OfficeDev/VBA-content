---
title: TabStrip.TabFixedWidth Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 3db4e0d2-c97b-a75b-3af6-b1678a1d5116
ms.date: 06/08/2017
---


# TabStrip.TabFixedWidth Property (Outlook Forms Script)

Returns or sets a  **Single** that represents the width in points of the tabs on a **[TabStrip](tabstrip-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **TabFixedWidth**

 _expression_A variable that represents a  **TabStrip** object.


## Remarks

If the value is 0, tab widths are automatically adjusted so that each tab is wide enough to accommodate its contents and each row of tabs spans the width of the control.

If the value is greater than 0, all tabs have an identical width as specified by this property.

The minimum size is 4 points.


