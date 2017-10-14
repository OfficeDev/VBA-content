---
title: WdHebSpellStart Enumeration (Word)
ms.prod: word
api_name:
- Word.WdHebSpellStart
ms.assetid: 9d0ca1f9-6bd6-08f1-fed9-71eb34ebc9ca
ms.date: 06/08/2017
---


# WdHebSpellStart Enumeration (Word)

Specifies which rules the Hebrew spelling checker will follow.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **wdFullScript**|0|The spelling checker follows rules for the conventional script required by the Hebrew Language Academy for writing text without diacritics.|
| **wdMixedAuthorizedScript**|3|The spelling checker follows rules for full and partial script, but highlights as potential mistakes any spelling variations not permitted within either system and any completely unrecognized words.|
| **wdMixedScript**|2|The spelling checker follows rules for full and partial script and allows non-conventional spelling variations. Only completely unrecognized words are highlighted as potential mistakes.|
| **wdPartialScript**|1|The spelling checker follows rules for the traditional script used only for text with diacritics.|

