---
title: Range.GetSpellingSuggestions Method (Word)
keywords: vbawd10.chm157155537
f1_keywords:
- vbawd10.chm157155537
ms.prod: word
api_name:
- Word.Range.GetSpellingSuggestions
ms.assetid: 5ab65e3e-65d8-4e49-2874-609b1974888e
ms.date: 06/08/2017
---


# Range.GetSpellingSuggestions Method (Word)

Returns a  **SpellingSuggestions** collection that represents the words suggested as spelling replacements for the first word in the specified range.


## Syntax

 _expression_ . **GetSpellingSuggestions**( **_CustomDictionary_** , **_IgnoreUppercase_** , **_MainDictionary_** , **_SuggestionMode_** , **_CustomDictionary2_** , **_CustomDictionary3_** , **_CustomDictionary4_** , **_CustomDictionary5_** , **_CustomDictionary6_** , **_CustomDictionary7_** , **_CustomDictionary8_** , **_CustomDictionary9_** , **_CustomDictionary10_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CustomDictionary_|Optional| **Variant**|Either an expression that returns a  **Dictionary** object or the file name of the custom dictionary.|
| _IgnoreUppercase_|Optional| **Variant**| **True** to ignore words in all uppercase letters. If this argument is omitted, the current value of the **IgnoreUppercase** property is used.|
| _MainDictionary_|Optional| **Variant**|Either an expression that returns a  **Dictionary** object or the file name of the main dictionary. If you don't specify a main dictionary, Microsoft Word uses the main dictionary that corresponds to the language formatting of the first word in the range.|
| _SuggestionMode_|Optional| **Variant**|Specifies the way Word makes spelling suggestions. Can be one of the following  **WdSpellingWordType** constants. The default value is **wdSpellword** .|
| _CustomDictionary2 ? CustomDictionary10_|Optional| **Variant**|Either an expression that returns a  **Dictionary** object or the file name of an additional custom dictionary. You can specify as many as nine additional dictionaries.|

### Return Value

SpellingSuggestions


## Remarks

If the word is spelled correctly, the  **Count** property of the **SpellingSuggestions** object returns 0 (zero).


## Example

This example looks for the alternate spelling suggestions for the first word in the selected range. If there are suggestions, the example runs a spelling check on the selection.


```vb
If Selection.Range.GetSpellingSuggestions.Count = 0 Then 
 Msgbox "No suggestions." 
Else 
 Selection.Range.CheckSpelling 
End If
```


## See also


#### Concepts


[Range Object](range-object-word.md)

