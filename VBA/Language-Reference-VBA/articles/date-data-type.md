---
title: Date Data Type
keywords: vblr6.chm1011011
f1_keywords:
- vblr6.chm1011011
ms.prod: office
ms.assetid: 728428b8-006d-aa0f-2532-f5154b1c56a4
ms.date: 06/08/2017
---


# Date Data Type

[Date variables](vbe-glossary.md) are stored as IEEE 64-bit (8-byte) floating-point numbers that represent dates ranging from 1 January 100 to 31 December 9999 and times from 0:00:00 to 23:59:59. Any recognizable literal date values can be assigned to **Date** variables.[Date literals](vbe-glossary.md) must be enclosed within number signs ( **#** ), for example, `#January 1, 1993#` or or `#1 Jan 93#`.

 **Date** variables display dates according to the short date format recognized by your computer. Times display according to the time format (either 12-hour or 24-hour) recognized by your computer.

When other [numeric types](vbe-glossary.md) are converted to **Date**, values to the left of the decimal represent date information while values to the right of the decimal represent time. Midnight is 0 and midday is 0.5. Negative whole numbers represent dates before 30 December 1899.


