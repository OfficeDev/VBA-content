---
title: MultiPage.TabOrientation Property (Outlook Forms Script)
keywords: olfm10.chm2002030
f1_keywords:
- olfm10.chm2002030
ms.prod: outlook
ms.assetid: 99a1d7ae-42b4-933c-2331-8b1c02550da6
ms.date: 06/08/2017
---


# MultiPage.TabOrientation Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the location of the tabs on a **[MultiPage](multipage-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **TabOrientation**

 _expression_A variable that represents a  **MultiPage** object.


## Remarks

The settings for  **TabOrientation** are:



|**Value**|**Description**|
|:-----|:-----|
|0|The tabs appear at the top of the control (default).|
|1|The tabs appear at the bottom of the control.|
|2|The tabs appear at the left side of the control.|
|3|The tabs appear at the right side of the control.|
If you use TrueType fonts, the text rotates when the  **TabOrientation** property is set to 2 or 3. If you use bitmapped fonts, the text does not rotate.


