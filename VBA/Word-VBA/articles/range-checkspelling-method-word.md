---
title: Range.CheckSpelling Method (Word)
keywords: vbawd10.chm157155533
f1_keywords:
- vbawd10.chm157155533
ms.prod: word
api_name:
- Word.Range.CheckSpelling
ms.assetid: 41873962-8cac-84a4-4e01-712985513cd4
ms.date: 06/08/2017
---


# Range.CheckSpelling Method (Word)

Begins a spelling check for the specified document or range.


## Syntax

 _expression_ . **CheckSpelling**( **_CustomDictionary_** , **_IgnoreUppercase_** , **_AlwaysSuggest_** , **_CustomDictionary2_** , **_CustomDictionary3_** , **_CustomDictionary4_** , **_CustomDictionary5_** , **_CustomDictionary6_** , **_CustomDictionary7_** , **_CustomDictionary8_** , **_CustomDictionary9_** , **_CustomDictionary10_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CustomDictionary_|Optional| **Variant**|Either an expression that returns a  **Dictionary** object or the file name of the custom dictionary.|
| _IgnoreUppercase_|Optional| **Variant**| **True** if capitalization is ignored. If this argument is omitted, the current value of the **IgnoreUppercase** property is used.|
| _AlwaysSuggest_|Optional| **Variant**| **True** for Microsoft Word to always suggest alternative spellings. If this argument is omitted, the current value of the **SuggestSpellingCorrections** property is used.|
| _CustomDictionary2_|Optional||Either an expression that returns a  **Dictionary** object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.|
| _CustomDictionary3_|Optional||Either an expression that returns a  **Dictionary** object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.|
| _CustomDictionary4_|Optional||Either an expression that returns a  **Dictionary** object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.|
| _CustomDictionary5_|Optional||Either an expression that returns a  **Dictionary** object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.|
| _CustomDictionary6_|Optional||Either an expression that returns a  **Dictionary** object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.|
| _CustomDictionary7_|Optional||Either an expression that returns a  **Dictionary** object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.|
| _CustomDictionary8_|Optional||Either an expression that returns a  **Dictionary** object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.|
| _CustomDictionary9_|Optional||Either an expression that returns a  **Dictionary** object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.|
| _CustomDictionary10_|Optional||Either an expression that returns a  **Dictionary** object or the file name of the custom dictionary. You can specify as many as nine additional dictionaries.|

## Remarks

If the document or range contains errors, this method displays the  **Spelling and Grammar** dialog box, with the **Check grammar** check box cleared.


## Example

This example begins a spelling check on all available stories of the active document.


```vb
Set range2 = Documents("MyDocument.doc").Sections(2).Range 
range2.CheckSpelling IgnoreUpperCase:=False, _ 
 CustomDictionary:="MyWork.Dic", _ 
 CustomDictionary2:="MyTechnical.Dic"
```


## See also


#### Concepts


[Range Object](range-object-word.md)

