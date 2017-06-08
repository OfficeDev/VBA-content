---
title: WdStylisticSet Enumeration (Word)
ms.prod: word
api_name:
- Word.WdStylisticSet
ms.assetid: e67291a0-5193-db3c-68da-3e3576da75c1
ms.date: 06/08/2017
---


# WdStylisticSet Enumeration (Word)

Specifies the stylistic set to apply to the font.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **wdStylisticSet01**|1|First stylistic set for the specified font.|
| **wdStylisticSet02**|2|Second stylistic set for the specified font.|
| **wdStylisticSet03**|4|Third stylistic set for the specified font.|
| **wdStylisticSet04**|8|Fourth stylistic set for the specified font.|
| **wdStylisticSet05**|16|Fifth stylistic set for the specified font.|
| **wdStylisticSet06**|32|Sixth stylistic set for the specified font.|
| **wdStylisticSet07**|64|Seventh stylistic set for the specified font.|
| **wdStylisticSet08**|128|Eighth stylistic set for the specified font.|
| **wdStylisticSet09**|256|Ninth stylistic set for the specified font.|
| **wdStylisticSet10**|512|Tenth stylistic set for the specified font.|
| **wdStylisticSet11**|1024|Eleventh stylistic set for the specified font.|
| **wdStylisticSet12**|2048|Twelfth stylistic set for the specified font.|
| **wdStylisticSet13**|4096|Thirteenth stylistic set for the specified font.|
| **wdStylisticSet14**|8192|Fourtheenth stylistic set for the specified font.|
| **wdStylisticSet15**|16384|Fifthteenth stylistic set for the specified font.|
| **wdStylisticSet16**|32768|Sixteenth stylistic set for the specified font.|
| **wdStylisticSet17**|65536|Seventeenth stylistic set for the specified font.|
| **wdStylisticSet18**|131072|Eighteenth stylistic set for the specified font.|
| **wdStylisticSet19**|262144|Nineteenth stylistic set for the specified font.|
| **wdStylisticSet20**|524288|Twentieth stylistic set for the specified font.|
| **wdStylisticSetDefault**|0|Default stylistic set for the specified font.|

## Remarks

Some OpenType fonts provide stylistic sets. A stylistic set defines a set of characters within the font that are intended to be used together, usually for the purpose of visual harmony, such as in headings. 20 stylistic sets are possible per font. 


 **Note**  Not all OpenType fonts provide stylistic sets. Setting a font's **[ StylisticSet](font-stylisticset-property-word.md)** property to a WdStylisticSet constant that is not provided by the font has no effect.


