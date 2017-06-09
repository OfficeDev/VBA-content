---
title: MultiPage.TabFixedWidth Property (Outlook Forms Script)
keywords: olfm10.chm2002000
f1_keywords:
- olfm10.chm2002000
ms.prod: outlook
ms.assetid: 932c2b27-97b7-adda-4ac5-3da64716f370
ms.date: 06/08/2017
---


# MultiPage.TabFixedWidth Property (Outlook Forms Script)

Returns or sets a  **Single** that represents the width in points of the tabs on a **[MultiPage](multipage-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **TabFixedWidth**

 _expression_A variable that represents a  **MultiPage** object.


## Remarks

If the value is 0, tab widths are automatically adjusted so that each tab is wide enough to accommodate its contents and each row of tabs spans the width of the control.

If the value is greater than 0, all tabs have an identical width as specified by this property.

The minimum size is 4 points.


