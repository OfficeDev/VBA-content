---
title: WdLigatures Enumeration (Word)
ms.prod: word
api_name:
- Word.WdLigatures
ms.assetid: 7441f3c4-a5cc-7ec4-cc57-2b1b0e05eb35
ms.date: 06/08/2017
---


# WdLigatures Enumeration (Word)

Specifies the type of ligatures applied to a font. 



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **wdLigaturesAll**|15|Applies all types of ligatures to the font.|
| **wdLigaturesContextual**|2|Applies contextual ligatures to the font. Contextual ligatures are often designed to enhance readability, but may also be solely ornamental. Contextual ligatures may also be contextual alternates.|
| **wdLigaturesContextualDiscretional**|10|Applies contextual and discretional ligatures to the font.|
| **wdLigaturesContextualHistorical**|6|Applies contextual and historical ligatures to the font.|
| **wdLigaturesContextualHistoricalDiscretional**|14|Applies contextual, historical, and discretional ligatures to a font.|
| **wdLigaturesDiscretional**|8|Applies discretional ligatures to the font. Discretional ligatures are most often designed to be ornamental at the discretion of the type developer.|
| **wdLigaturesHistorical**|4|Applies historical ligatures to the font. Historical ligatures are similar to standard ligatures in that they were originally intended to improve the readability of the font, but may look archaic to the modern reader.|
| **wdLigaturesHistoricalDiscretional**|12|Applies historical and discretional ligatures to the font.|
| **wdLigaturesNone**|0|Does not apply any ligatures to the font.|
| **wdLigaturesStandard**|1|Applies standard ligatures to the font. Standard ligatures are designed to enhance readability. Standard ligatures in Latin languages include "fi", "fl", and "ff", for example.|
| **wdLigaturesStandardContextual**|3|Applies standard and contextual ligatures to the font.|
| **wdLigaturesStandardContextualDiscretional**|11|Applies standard, contextual and discretional ligatures to the font.|
| **wdLigaturesStandardContextualHistorical**|7|Applies standard, contextual, and historical ligatures to the font.|
| **wdLigaturesStandardDiscretional**|9|Applies standard and discretional ligatures to the font.|
| **wdLigaturesStandardHistorical**|5|Applies standard and historical ligatures to the font.|
| **wdLigaturesStandardHistoricalDiscretional**|13|Applies standard historical and discretional ligatures to the font.|

## Remarks

A glyph is a visual representation of a character. Ligatures are two or more glyphs that are represented by what appears to the reader as a single character in order to create more readable or attractive text. Use the [Font.Ligatures Property (Word)](font-ligatures-property-word.md) property to specify the ligatures to apply to a font in Word.


 **Note**  The order of preference when a combination of ligature types are applied differs by font and is not controlled by the Word application.


