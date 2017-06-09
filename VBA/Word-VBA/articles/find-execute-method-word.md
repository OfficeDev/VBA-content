---
title: Find.Execute Method (Word)
keywords: vbawd10.chm162529724
f1_keywords:
- vbawd10.chm162529724
ms.prod: word
api_name:
- Word.Find.Execute
ms.assetid: 3b607955-0e82-aa13-dad1-7a5069a57b9d
ms.date: 06/08/2017
---


# Find.Execute Method (Word)

Runs the specified find operation. Returns  **True** if the find operation is successful. **Boolean** .


## Syntax

 _expression_ . **Execute**( **_FindText_** , **_MatchCase_** , **_MatchWholeWord_** , **_MatchWildcards_** , **_MatchSoundsLike_** , **_MatchAllWordForms_** , **_Forward_** , **_Wrap_** , **_Format_** , **_ReplaceWith_** , **_Replace_** , **_MatchKashida_** , **_MatchDiacritics_** , **_MatchAlefHamza_** , **_MatchControl_** )

 _expression_ Required. A variable that represents a **[Find](find-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FindText_|Optional| **Variant**|The text to be searched for. Use an empty string ("") to search for formatting only. You can search for special characters by specifying appropriate character codes. For example, "^p" corresponds to a paragraph mark and "^t" corresponds to a tab character.|
| _MatchCase_|Optional| **Variant**| **True** to specify that the find text be case sensitive. Corresponds to the **Match case** check box in the **Find and Replace** dialog box ( **Edit** menu).|
| _MatchWholeWord_|Optional| **Variant**| **True** to have the find operation locate only entire words, not text that is part of a larger word. Corresponds to the **Find whole words only** check box in the **Find and Replace** dialog box.|
| _MatchWildcards_|Optional| **Variant**| **True** to have the find text be a special search operator. Corresponds to the **Use wildcards** check box in the **Find and Replace** dialog box.|
| _MatchSoundsLike_|Optional| **Variant**| **True** to have the find operation locate words that sound similar to the find text. Corresponds to the **Sounds like** check box in the **Find and Replace** dialog box.|
| _MatchAllWordForms_|Optional| **Variant**| **True** to have the find operation locate all forms of the find text (for example, "sit" locates "sitting" and "sat"). Corresponds to the **Find all word forms** check box in the **Find and Replace** dialog box.|
| _Forward_|Optional| **Variant**| **True** to search forward (toward the end of the document).|
| _Wrap_|Optional| **Variant**|Controls what happens if the search begins at a point other than the beginning of the document and the end of the document is reached (or vice versa if Forward is set to  **False** ). This argument also controls what happens if there is a selection or range and the search text is not found in the selection or range. Can be one of the **WdFindWrap** constants.|
| _Format_|Optional| **Variant**| **True** to have the find operation locate formatting in addition to, or instead of, the find text.|
| _ReplaceWith_|Optional| **Variant**|The replacement text. To delete the text specified by the Find argument, use an empty string (""). You specify special characters and advanced search criteria just as you do for the Find argument. To specify a graphic object or other nontext item as the replacement, move the item to the Clipboard and specify "^c" for ReplaceWith.|
| _Replace_|Optional| **Variant**|Specifies how many replacements are to be made: one, all, or none. Can be any  **[WdReplace](http://msdn.microsoft.com/library/e7e8b8c3-e862-5fe6-ee56-b054263a4402.md.aspx)** constant.|
| _MatchKashida_|Optional| **Variant**| **True** if find operations match text with matching kashidas in an Arabic-language document. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|
| _MatchDiacritics_|Optional| **Variant**| **True** if find operations match text with matching diacritics in a right-to-left language document. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|
| _MatchAlefHamza_|Optional| **Variant**| **True** if find operations match text with matching alef hamzas in an Arabic-language document. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|
| _MatchControl_|Optional| **Variant**| **True** if find operations match text with matching bidirectional control characters in a right-to-left language document. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|
| _MatchPrefix_|Optional| **Variant**| **True** to match words beginning with the search string. Corresponds to the **Match prefix** check box in the **Find and Replace** dialog box.|
| _MatchSuffix_|Optional| **Variant**| **True** to match words ending with the search string. Corresponds to the **Match suffix** check box in the **Find and Replace** dialog box.|
| _MatchPhrase_|Optional| **Variant**| **True** ignores all white space and control characters between words.|
| _IgnoreSpace_|Optional| **Variant**| **True** ignores all white space between words. Corresponds to the **Ignore white-space characters** check box in the **Find and Replace** dialog box.|
| _IgnorePunct_|Optional| **Variant**| **True** ignores all punctuation characters between words. Corresponds to the **Ignore punctuation** check box in the **Find and Replace** dialog box.|

### Return Value

Boolean


## Remarks

If  **MatchWildcards** is **True** , you can specify wildcard characters and other advanced search criteria for the FindText argument. For example, "*(ing)" finds any word that ends in "ing".

To search for a symbol character, type a caret (^), a zero (0), and then the symbol's character code. For example, "^0151" corresponds to an em dash (?).

Unless otherwise specified, replacement text inherits the formatting of the text it replaces in the document. For example, if you replace the string "abc" with "xyz", occurrences of "abc" with bold formatting are replaced with the string "xyz" with bold formatting.

Also, if  **MatchCase** is **False** , occurrences of the search text that are uppercase will be replaced with an uppercase version of the replacement text, regardless of the case of the replacement text. Using the previous example, occurrences of "ABC" are replaced with "XYZ".


## Example

This example finds and selects the next occurrence of the word "library".


```vb
With Selection.Find 
    .ClearFormatting 
    .MatchWholeWord = True 
    .MatchCase = False 
    .Execute FindText:="library" 
End With
```

This example finds all occurrences of the word "hi" in the active document and replaces each occurrence with "hello".




```vb
Set myRange = ActiveDocument.Content 
myRange.Find.Execute FindText:="hi", _ 
    ReplaceWith:="hello", Replace:=wdReplaceAll
```


## See also


#### Concepts


[Find Object](find-object-word.md)

