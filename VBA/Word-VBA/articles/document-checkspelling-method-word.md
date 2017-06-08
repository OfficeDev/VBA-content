---
title: Document.CheckSpelling Method (Word)
keywords: vbawd10.chm158007428
f1_keywords:
- vbawd10.chm158007428
ms.prod: word
api_name:
- Word.Document.CheckSpelling
ms.assetid: a61a9c8b-0dee-f6e4-cefc-daca612c99c1
ms.date: 06/08/2017
---


# Document.CheckSpelling Method (Word)

Begins a spelling check for the specified document or range. .


## Syntax

 _expression_ . **CheckSpelling**( **_CustomDictionary_** , **_IgnoreUppercase_** , **_AlwaysSuggest_** , **_CustomDictionary2_** , **_CustomDictionary3_** , **_CustomDictionary4_** , **_CustomDictionary5_** , **_CustomDictionary6_** , **_CustomDictionary7_** , **_CustomDictionary8_** , **_CustomDictionary9_** , **_CustomDictionary10_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _IgnoreUppercase_|Optional| **Variant**| **True** if capitalization is ignored. If this argument is omitted, the current value of the **IgnoreUppercase** property is used.|
| _AlwaysSuggest_|Optional| **Variant**| **True** for Microsoft Word to always suggest alternative spellings. If this argument is omitted, the current value of the **SuggestSpellingCorrections** property is used.|

## Remarks

If the document or range contains errors, this method displays the  **Spelling and Grammar** dialog box ( **Tools** menu), with the **Check grammar** check box cleared. For a document, this method checks all available stories (such as headers, footers, and text boxes).


## Example

The following example checks the spelling in the active document.


```vb
ActiveDocument.CheckSpelling
```


## See also


#### Concepts


[Document Object](document-object-word.md)

