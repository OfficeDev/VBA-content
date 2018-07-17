---
title: Find.Execute2007 Method (Word)
keywords: vbawd10.chm162529727
f1_keywords:
- vbawd10.chm162529727
ms.prod: word
api_name:
- Word.Find.Execute2007
ms.assetid: 441de4b6-882c-e950-cafe-ee4463ef1007
ms.date: 06/08/2017
---


# Find.Execute2007 Method (Word)

Runs the specified find operation. Returns  **True** if the find operation is successful.


## Syntax

 _expression_ . **Execute2007**( **_FindText_** , **_MatchCase_** , **_MatchWholeWord_** , **_MatchWildcards_** , **_MatchSoundsLike_** , **_MatchAllWordForms_** , **_Forward_** , **_Wrap_** , **_Format_** , **_ReplaceWith_** , **_Replace_** , **_MatchKashida_** , **_MatchDiacritics_** , **_MatchAlefHamza_** , **_MatchControl_** , **_MatchPrefix_** , **_MatchSuffix_** , **_MatchPhrase_** , **_IgnoreSpace_** , **_IgnorePunct_** )

 _expression_ A variable that represents a **[Find](find-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FindText_|Optional| **Variant**|The text to be searched for. Use an empty string ("") to search for formatting only. You can search for special characters by specifying appropriate character codes. For example, "^p" corresponds to a paragraph mark and "^t" corresponds to a tab character.|
| _MatchCase_|Optional| **Variant**| **True** to specify that the text to find should be case-sensitive. Corresponds to the **Match case** check box in the **Find and Replace** dialog box (on the **Home** tab in the **Editing** group).|
| _MatchWholeWord_|Optional| **Variant**| **True** to find only entire words, not text that is part of a larger word. Corresponds to the **Find whole words only** check box in the **Find and Replace** dialog box.|
| _MatchWildcards_|Optional| **Variant**| **True** to use wildcard search operators in the text to find. Corresponds to the **Use wildcards** check box in the **Find and Replace** dialog box.|
| _MatchSoundsLike_|Optional| **Variant**| **True** to locate words that sound similar to the text to find. Corresponds to the **Sounds like** check box in the **Find and Replace** dialog box.|
| _MatchAllWordForms_|Optional| **Variant**| **True** to locate all forms of the text to find (for example, "sit" locates "sitting" and "sat"). Corresponds to the **Find all word forms** check box in the **Find and Replace** dialog box.|
| _Forward_|Optional| **Variant**| **True** to search forward (toward the end of the document).|
| _Wrap_|Optional| **Variant**|One of the [WdFindWrap](wdfindwrap-enumeration-word.md) constants that controls what happens if the search begins at a point other than the beginning of the document and the end of the document is reached (or vice versa if Forward is set to **False** ). This argument also controls what happens if there is a selection or range and the search text is not found in the selection or range.|
| _Format_|Optional| **Variant**| **True** to locate formatting in addition to, or instead of, the text to find.|
| _ReplaceWith_|Optional| **Variant**|The replacement text. To delete the text specified by the Find argument, use an empty string (""). You specify special characters and advanced search criteria just as you do for the Find argument. To specify a graphic object or other nontext item as the replacement, move the item to the Clipboard and specify "^c" for ReplaceWith.|
| _Replace_|Optional| **Variant**|One of the [WdReplace](wdreplace-enumeration-word.md) constants that specifies how many replacements are to be made: one, all, or none.|
| _MatchKashida_|Optional| **Variant**| **True** to find matching kashidas in an Arabic-language document. This argument may not be available to you, depending on the language support (for example, U.S. English) that you have selected or installed.|
| _MatchDiacritics_|Optional| **Variant**| **True** to find matching diacritics in a right-to-left language document. This argument may not be available to you, depending on the language support (for example, U.S. English) that you have selected or installed.|
| _MatchAlefHamza_|Optional| **Variant**| **True** to find matching alef hamzas in an Arabic-language document. This argument may not be available to you, depending on the language support (for example, U.S. English) that you have selected or installed.|
| _MatchControl_|Optional| **Variant**| **True** to find matching bidirectional control characters in a right-to-left language document. This argument may not be available to you, depending on the language support (for example, U.S. English) that you have selected or installed.|
| _MatchPrefix_|Optional| **Variant**| **True** to find words that begin with the search string. Corresponds to the **Match prefix** check box in the **Find and Replace** dialog box.|
| _MatchSuffix_|Optional| **Variant**| **True** to find words that end with the search string. Corresponds to the **Match suffix** check box in the **Find and Replace** dialog box.|
| _MatchPhrase_|Optional| **Variant**| **True** to ignore all white space and control characters between words.|
| _IgnoreSpace_|Optional| **Variant**| **True** to ignore all white space between words. Corresponds to the **Ignore white-space characters** check box in the **Find and Replace** dialog box.|
| _IgnorePunct_|Optional| **Variant**| **True** to ignore all punctuation characters between words. Corresponds to the **Ignore punctuation** check box in the **Find and Replace** dialog box.|

### Return Value

A  **Boolean** value that indicates whether the find operation was successful.


## Remarks

If MatchWildcards is  **True** , you can specify wildcard characters and other advanced search criteria for the FindText argument. For example, "*(ing)" finds any word that ends in "ing".

To search for a symbol character, type a caret (^), a zero (0), and then the symbol's character code. For example, "^0151" corresponds to an em dash (?).

Unless otherwise specified, replacement text inherits the formatting of the text it replaces in the document. For example, if you replace the string "abc" with "xyz", occurrences of "abc" with bold formatting are replaced with the string "xyz" with bold formatting.

Also, if MatchCase is  **False** , occurrences of the search text that are uppercase will be replaced with an uppercase version of the replacement text, regardless of the case of the replacement text. Using the previous example, occurrences of "ABC" are replaced with "XYZ".


## Example

The following example finds and selects the next occurrence of the word "library".


```vb
With Selection.Find 
 .ClearFormatting 
 .MatchWholeWord = True 
 .MatchCase = False 
 .Execute2007 FindText:="library" 
End With
```

The following example finds all occurrences of the word "hi" in the active document and replaces each occurrence with "hello".




```vb
Set myRange = ActiveDocument.Content 
myRange.Find.Execute2007 FindText:="hi", _ 
 ReplaceWith:="hello", Replace:=wdReplaceAll
```


## See also


#### Concepts


[Find Object](find-object-word.md)

