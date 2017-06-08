---
title: StrConv Constants
keywords: vblr6.chm1012530
f1_keywords:
- vblr6.chm1012530
ms.prod: office
ms.assetid: bac42216-f443-439a-d346-f74da2d98edd
ms.date: 06/08/2017
---


# StrConv Constants

The following [constants](vbe-glossary.md) can be used anywhere in your code in place of the actual values:



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbUpperCase**|1|Converts the string to uppercase characters.|
|**vbLowerCase**|2|Converts the string to lowercase characters.|
|**vbProperCase**|3|Converts the first letter of every word in string to uppercase.|
|**vbWide**|4|Converts narrow (single-byte) characters in string to wide (double-byte) characters. Applies to East Asia [locales](vbe-glossary.md).|
|**vbNarrow**|8|Converts wide (double-byte) characters in string to narrow (single-byte) characters. Applies to East Asia locales.|
|**vbKatakana**|16|Converts Hiragana characters in string to Katakana characters. Applies to Japan only.|
|**vbHiragana**|32|Converts Katakana characters in string to Hiragana characters. Applies to Japan only.|
|**vbUnicode**|64|Converts the string to [Unicode](vbe-glossary.md) using the default code page of the system. (Not available on the Macintosh.)|
|**vbFromUnicode**|128|Converts the string from Unicode to the default code page of the system. (Not available on the Macintosh.)|

