---
title: Find.HitHighlight Method (Word)
keywords: vbawd10.chm162529725
f1_keywords:
- vbawd10.chm162529725
ms.prod: word
api_name:
- Word.Find.HitHighlight
ms.assetid: 11f6a7e5-7aba-a374-db39-327f6427364b
ms.date: 06/08/2017
---


# Find.HitHighlight Method (Word)

Highlights all found matches and returns a **Boolean** that represents whether matches were found.

## Syntax

_expression_. **HitHighlight** (**_FindText_**, **_HighlightColor_**, **_TextColor_**, **_MatchCase_**, **_MatchWholeWord_**, **_MatchPrefix_**, **_MatchSuffix_**, **_MatchPhrase_**, **_MatchWildcards_**, **_MatchSoundsLike_**, **_MatchAllWordForms_**, **_MatchByte_**, **_MatchFuzzy_**, **_MatchKashida_**, **_MatchDiacritics_**, **_MatchAlefHamza_**, **_MatchControl_**, **_IgnoreSpace_**, **_IgnorePunct_**, **_HanjaPhoneticHangul_**)

_expression_ An expression that returns a **Find** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FindText_|Required| **Variant**|Specifies the text to find. Use an empty string ("") to search for formatting only. You can search for special characters by specifying appropriate character codes. For example, `"^p"` corresponds to a paragraph mark and `"^t"` corresponds to a tab character.|
| _HighlightColor_|Optional| **Variant**|Specifies the highlight color for the text. Can be any RGB color or one of the **[WdColor](wdcolor-enumeration-word.md)** constants.|
| _TextColor_|Optional| **Variant**|Specifies the color of the text. Can be any RGB color or one of the **[WdColor](wdcolor-enumeration-word.md)** constants.|
| _MatchCase_|Optional| **Variant**|**True** to specify that the find text be case-sensitive. Corresponds to the **Match case** check box in the **Find and Replace** dialog box.|
| _MatchWholeWord_|Optional| **Variant**|**True** to have the find operation locate only entire words, not text that is part of a larger word. Corresponds to the **Find whole words only** check box in the **Find and Replace** dialog box.|
| _MatchPrefix_|Optional| **Variant**|**True** to match words beginning with the search string. Corresponds to the **Match prefix** check box in the **Find and Replace** dialog box.|
| _MatchSuffix_|Optional| **Variant**|**True** to match words ending with the search string. Corresponds to the **Match suffix** check box in the **Find and Replace** dialog box.|
| _MatchPhrase_|Optional| **Variant**|**True** ignores all white space and control characters between words.|
| _MatchWildcards_|Optional| **Variant**|**True** to have the find text be a special search operator. Corresponds to the **Use wildcards** check box in the **Find and Replace** dialog box.|
| _MatchSoundsLike_|Optional| **Variant**|**True** to have the find operation locate words that sound similar to the find text. Corresponds to the **Sounds like** check box in the **Find and Replace** dialog box.|
| _MatchAllWordForms_|Optional| **Variant**|**True** to have the find operation locate all forms of the find text (for example, "sit" locates "sitting" and "sat"). Corresponds to the **Find all word forms** check box in the **Find and Replace** dialog box.|
| _MatchByte_|Optional| **Variant**|**True** to distinguish between full-width and half-width letters or characters during a search.|
| _MatchFuzzy_|Optional| **Variant**|**True** to use the nonspecific search options for Japanese text during a search. Read/write.|
| _MatchKashida_|Optional| **Variant**|**True** if find operations match text with matching kashidas in an Arabic-language document. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|
| _MatchDiacritics_|Optional| **Variant**|**True** if find operations match text with matching diacritics in a right-to-left language document. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|
| _MatchAlefHamza_|Optional| **Variant**|**True** if find operations match text with matching alef hamzas in an Arabic-language document. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|
| _MatchControl_|Optional| **Variant**|**True** if find operations match text with matching bidirectional control characters in a right-to-left language document. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|
| _IgnoreSpace_|Optional| **Variant**|**True** ignores all white space between words. Corresponds to the **Ignore white-space characters** check box in the **Find and Replace** dialog box.|
| _IgnorePunct_|Optional| **Variant**|**True** ignores all punctuation characters between words. Corresponds to the **Ignore punctuation** check box in the **Find and Replace** dialog box.|
| _HanjaPhoneticHangul_|Optional| **Variant**|**True** ignores phonetic hangul and hanja characters. Available only if you have support for Korean languages.|

### Return value

Boolean

## Remarks

The **HitHighlight** method applies primarily to search highlighting in Office Outlook when Word is specified as the e-mail editor. However, you can use this method on documents inside Word if you want to highlight found text. Otherwise, use the **[Execute](find-execute-method-word.md)** method.

## See also

#### Concepts

- [Find Object](find-object-word.md)

