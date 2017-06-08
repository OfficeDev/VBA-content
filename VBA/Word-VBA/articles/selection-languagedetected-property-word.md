---
title: Selection.LanguageDetected Property (Word)
keywords: vbawd10.chm158663663
f1_keywords:
- vbawd10.chm158663663
ms.prod: word
api_name:
- Word.Selection.LanguageDetected
ms.assetid: 289e6a01-1945-a17f-f6a0-e472cfa263eb
ms.date: 06/08/2017
---


# Selection.LanguageDetected Property (Word)

Returns or sets a  **Boolean** that specifies whether Microsoft Word has detected the language of the selected text.


## Syntax

 _expression_ . **LanguageDetected**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

Check the  **LanguageID** property for the results of any previous language detection.

The  **LanguageDetected** property is set to **True** when the **DetectLanguage** method is called. To reevaluate the language of the specified text, you must first set the **LanguageDetected** property to **False** .


## Example

This example checks the active document to determine the language it's written in and then displays the result.


```vb
With Selection 
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


[Selection Object](selection-object-word.md)

