---
title: Application.Language Property (Word)
keywords: vbawd10.chm158335367
f1_keywords:
- vbawd10.chm158335367
ms.prod: word
api_name:
- Word.Application.Language
ms.assetid: b24f0861-1f7a-ecd9-7084-39c65df4fcc3
ms.date: 06/08/2017
---


# Application.Language Property (Word)

Returns an  **MsoLanguageID** constant that represents the language selected for the Microsoft Word user interface.


## Syntax

 _expression_ . **Language**

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

The value of this property is the same as the value returned by the following expression:


```vb
Application.LanguageSettings _ 
 .LanguageID(msoLanguageIDUI)
```


## Example

This example displays a message stating whether the language selected for the Microsoft Word user interface is U.S. English.


```vb
Sub LangSetting() 
 If Application.Language = msoLanguageIDEnglishUS Then 
 MsgBox "The user interface language is U.S. English." 
 Else 
 MsgBox "The user interface language is not U.S. English." 
 End If 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

