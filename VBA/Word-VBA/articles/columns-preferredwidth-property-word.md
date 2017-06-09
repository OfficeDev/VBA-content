---
title: Columns.PreferredWidth Property (Word)
keywords: vbawd10.chm155910249
f1_keywords:
- vbawd10.chm155910249
ms.prod: word
api_name:
- Word.Columns.PreferredWidth
ms.assetid: 72a64aaa-0c53-2e61-9c33-fb10436823e9
ms.date: 06/08/2017
---


# Columns.PreferredWidth Property (Word)

Returns or sets the preferred width (in points or as a percentage of the window width) for the specified columns. Read/write  **Single** .


## Syntax

 _expression_ . **PreferredWidth**

 _expression_ Required. An expression that returns a **[Columns](columns-object-word.md)** collection.


## Remarks

If the  **[PreferredWidthType](columns-preferredwidthtype-property-word.md)** property is set to **wdPreferredWidthPoints** , the **PreferredWidth** property returns or sets the width in points. If the **PreferredWidthType** property is set to **wdPreferredWidthPercent** , the **PreferredWidth** property returns or sets the width as a percentage of the window width.


## See also


#### Concepts


[Columns Collection Object](columns-object-word.md)

