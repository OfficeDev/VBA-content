---
title: Line too long
keywords: vblr6.chm1011210
f1_keywords:
- vblr6.chm1011210
ms.prod: office
ms.assetid: e63efae7-1383-2da7-8416-94d104e4abd4
ms.date: 06/08/2017
---


# Line too long

A physical line of Visual Basic code can contain up to 1023 characters. This error has the following cause and solution:



- A line contains too many characters. You can create a longer logical line by joining physical lines with a [line-continuation character](vbe-glossary.md), a space followed by an underscore ( _). Up to 25 physical lines can be joined with line-continuation characters to form a single logical line, or 24 consecutive line-continuation characters. Thus, a logical line could potentially contain a total of 10,230 characters. Beyond that, you must break the line into individual statements or assign some [expressions](vbe-glossary.md) to intermediate[variables](vbe-glossary.md).
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

