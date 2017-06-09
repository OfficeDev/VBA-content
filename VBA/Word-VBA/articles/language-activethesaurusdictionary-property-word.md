---
title: Language.ActiveThesaurusDictionary Property (Word)
keywords: vbawd10.chm158138384
f1_keywords:
- vbawd10.chm158138384
ms.prod: word
api_name:
- Word.Language.ActiveThesaurusDictionary
ms.assetid: 2fedc56e-e694-56a7-0ce9-7ff45c6cbed1
ms.date: 06/08/2017
---


# Language.ActiveThesaurusDictionary Property (Word)

Returns a  **[Dictionary](dictionary-object-word.md)** object that represents the active thesaurus dictionary for the specified language.


## Syntax

 _expression_ . **ActiveThesaurusDictionary**

 _expression_ An expression that returns a **[Language](language-object-word.md)** object.


## Remarks

If there is no thesaurus dictionary installed for the specified language, this property returns  **Nothing** .


## Example

This example returns the full path and file name of the active thesaurus dictionary.


```vb
Dim lngLanguage As Long 
Dim dicThesaurus As Dictionary 
 
lngLanguage = Selection.LanguageID 
Set dicThesaurus = Languages(lngLanguage).ActiveThesaurusDictionary 
If dicThesaurus Is Nothing Then 
 MsgBox "No thesaurus dictionary installed!" 
Else 
 MsgBox dicThesaurus.Path &; Application.PathSeparator _ 
 &; dicThesaurus.Name 
End If 

```


## See also


#### Concepts


[Language Object](language-object-word.md)

