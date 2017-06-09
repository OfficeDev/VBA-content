---
title: Range.LanguageDetected Property (Word)
keywords: vbawd10.chm157155591
f1_keywords:
- vbawd10.chm157155591
ms.prod: word
api_name:
- Word.Range.LanguageDetected
ms.assetid: dfe307e5-ad87-9a6b-ecbe-521c6354b349
ms.date: 06/08/2017
---


# Range.LanguageDetected Property (Word)

Returns or sets a value that specifies whether Microsoft Word has detected the language of the specified text. Read/write  **Boolean** .


## Syntax

 _expression_ . **LanguageDetected**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

Check the  **LanguageID** property for the results of any previous language detection.

The  **LanguageDetected** property is set to **True** when the **DetectLanguage** method is called. To reevaluate the language of the specified text, you must first set the **LanguageDetected** property to **False** .


## Example

This example checks the active document to determine the language it's written in and then displays the result.


```vb
With ActiveDocument.Range 
 If .LanguageDetected = True Then 
 x = MsgBox("This document has already " _ 
 &; "been checked. Do you want to check " _ 
 &; "it again?", vbYesNo) 
 If x = vbYes Then 
 .LanguageDetected = False 
 .DetectLanguage 
 End If 
 Else 
 .DetectLanguage 
 End If 
 If .Range.LanguageID = wdEnglishUS Then 
 MsgBox "This is a U.S. English document." 
 Else 
 MsgBox "This is not a U.S. English document." 
 End If 
End With
```


## See also


#### Concepts


[Range Object](range-object-word.md)

