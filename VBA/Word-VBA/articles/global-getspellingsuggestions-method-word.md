---
title: Global.GetSpellingSuggestions Method (Word)
keywords: vbawd10.chm163119431
f1_keywords:
- vbawd10.chm163119431
ms.prod: word
api_name:
- Word.Global.GetSpellingSuggestions
ms.assetid: 1539a24d-1548-d330-b90b-98d118b999c4
ms.date: 06/08/2017
---


# Global.GetSpellingSuggestions Method (Word)

Returns a  **[SpellingSuggestions](spellingsuggestions-object-word.md)** collection that represents the words suggested as spelling replacements for a given word.


## Syntax

 _expression_ . **GetSpellingSuggestions**( **_Word_** , **_CustomDictionary_** , **_IgnoreUppercase_** , **_MainDictionary_** , **_SuggestionMode_** , **_CustomDictionary2_** , **_CustomDictionary3_** , **_CustomDictionary4_** , **_CustomDictionary5_** , **_CustomDictionary6_** , **_CustomDictionary7_** , **_CustomDictionary8_** , **_CustomDictionary9_** , **_CustomDictionary10_** )

 _expression_ Required. A variable that represents a **[Global](global-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Word_|Required| **String**|The word whose spelling is to be checked.|
| _IgnoreUppercase_|Optional| **Variant**| **True** to ignore words in all uppercase letters. If this argument is omitted, the current value of the **[IgnoreUppercase](options-ignoreuppercase-property-word.md)** property is used.|
| _SuggestionMode_|Optional| **Variant**|Specifies the way Word makes spelling suggestions. Can be one of the following  **[WdSpellingWordType](wdspellingwordtype-enumeration-word.md)** constants: **wdAnagram** , **wdSpellword** , or **wdWildcard** . The default value is **WdSpellword** .|

## Remarks

If the word is spelled correctly, the  **Count** property of the **[SpellingSuggestions](spellingsuggestions-object-word.md)** object returns 0 (zero).


## Example

This example looks for the alternate spelling suggestions for the word "?ook." Suggestions include replacements for the ? wildcard character. Any suggested spellings are displayed in message boxes.


```vb
Sub DisplaySuggestions() 
 Dim sugList As SpellingSuggestions 
 Dim sug As SpellingSuggestion 
 Dim strSugList As String 
 Set sugList = GetSpellingSuggestions(Word:="lrok", _ 
 SuggestionMode:=wdSpellword) 
 If sugList.Count = 0 Then 
 MsgBox "No suggestions." 
 Else 
 For Each sug In sugList 
 strSugList = strSugList &; vbTab &; sug.Name &; vbLf 
 Next sug 
 MsgBox "The suggestions for this word are: " _ 
 &; vbLf &; strSugList 
 End If 
End Sub
```


## See also


#### Concepts


[Global Object](global-object-word.md)

