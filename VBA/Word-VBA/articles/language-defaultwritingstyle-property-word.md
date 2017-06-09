---
title: Language.DefaultWritingStyle Property (Word)
keywords: vbawd10.chm158138385
f1_keywords:
- vbawd10.chm158138385
ms.prod: word
api_name:
- Word.Language.DefaultWritingStyle
ms.assetid: 89eae276-8439-35d1-19bf-92c8ba69575c
ms.date: 06/08/2017
---


# Language.DefaultWritingStyle Property (Word)

Returns or sets the default writing style used by the grammar checker for the specified language. Read/write  **String** .


## Syntax

 _expression_ . **DefaultWritingStyle**

 _expression_ A variable that represents a **[Language](language-object-word.md)** object.


## Remarks

This property controls the global setting for the writing style. The name of the writing style is the localized name for the specified language. When setting this property, you must use the exact name found in the  **Writing style box** on the **Spelling &; Grammar** tab in the **Options** dialog box ( **Tools** menu).

The  **[ActiveWritingStyle](document-activewritingstyle-property-word.md)** property sets the writing style for each language in a document. The **ActiveWritingStyle** setting overrides the **DefaultWritingStyle** setting.


## Example

This example returns the default writing style in a message box.


```vb
Dim lngLanguage As Long 
 
lngLanguage = Selection.LanguageID 
Msgbox Languages(lngLanguage).DefaultWritingStyle
```

This example sets the writing style for U.S. English to Casual, and then it checks spelling and grammar in the active document.




```
Languages(wdEnglishUS).DefaultWritingStyle = "Casual" 
ActiveDocument.CheckGrammar
```


## See also


#### Concepts


[Language Object](language-object-word.md)

