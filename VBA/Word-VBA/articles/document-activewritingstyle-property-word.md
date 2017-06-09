---
title: Document.ActiveWritingStyle Property (Word)
keywords: vbawd10.chm158007386
f1_keywords:
- vbawd10.chm158007386
ms.prod: word
api_name:
- Word.Document.ActiveWritingStyle
ms.assetid: 035c0872-8c0b-c95f-dd0c-893982304e0f
ms.date: 06/08/2017
---


# Document.ActiveWritingStyle Property (Word)

Returns or sets the writing style for a specified language in the specified document. Read/write  **String** .


## Syntax

 _expression_ . **ActiveWritingStyle**( **_LanguageID_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _LanguageID_|Required| **Variant**|The language to set the writing style for in the specified document. Can be either a string or one of the following  **WdLanguageID** constants. Some of the **WdLanguageID** constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|

## Remarks

The  **WritingStyleList** property returns an array of the names of the available writing styles.


## Example

This example sets the writing style used for French, German, and U.S. English for the active document. You must have the grammar files installed for French, German, and U.S. English to run this example.


```vb
With ActiveDocument 
 .ActiveWritingStyle(wdFrench) = "Commercial" 
 .ActiveWritingStyle(wdGerman) = "Technisch/Wiss" 
 .ActiveWritingStyle(wdEnglishUS) = "Technical" 
End With
```

This example returns the writing style for the language of the selection.




```vb
Sub WhichLanguage() 
 Dim varLang As Variant 
 
 varLang = Selection.LanguageID 
 MsgBox ActiveDocument.ActiveWritingStyle(varLang) 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

