---
title: Range.DetectLanguage Method (Word)
keywords: vbawd10.chm157155531
f1_keywords:
- vbawd10.chm157155531
ms.prod: word
api_name:
- Word.Range.DetectLanguage
ms.assetid: 4b4149fa-011a-2489-8779-66d75897174f
ms.date: 06/08/2017
---


# Range.DetectLanguage Method (Word)

Analyzes the specified text to determine the language that it is written in.


## Syntax

 _expression_ . **DetectLanguage**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

The results of the  **DetectLanguage** method are stored in the **LanguageID** property on a character-by-character basis. To read the **[LanguageID](language-id-property-word.md)** property, you must first specify a selection or range of text.



When applied to a  **Document** object, the **DetectLanguage** method checks all available text in the document (headers, footers, text boxes, and so forth). If the specified text contains a partial sentence, the selection or range is extended to the end of the sentence.



If the  **DetectLanguage** method has already been applied to the specified text, the **LanguageDetected** property is set to **True** . To reevaulate the language of the specified text, you must first set the **[LanguageDetected](document-languagedetected-property-word.md)** property to **False** .



For more information about automatic language detection, see About automatic language detection .




## Example

This example checks the active document to determine the language it's written in and then displays the result.


```vb
With ActiveDocument 
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

