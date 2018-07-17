---
title: Style.LanguageID Property (Word)
keywords: vbawd10.chm153878540
f1_keywords:
- vbawd10.chm153878540
ms.prod: word
api_name:
- Word.Style.LanguageID
ms.assetid: 83c4bebe-4c8a-cd38-5083-4a227c09a07d
ms.date: 06/08/2017
---


# Style.LanguageID Property (Word)

Returns or sets a  **[WdLanguageID](wdlanguageid-enumeration-word.md)** constant that represents the language for the specified range. Read/write.


## Syntax

 _expression_ . **LanguageID**

 _expression_ An expression that represents a **[Style](style-object-word.md)** object.


## Remarks

Some of the  **WdLanguageID** constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example redefines the Title style to use the Spanish proofing tools. The new style description is then displayed in a message box.


```vb
ActiveDocument.Styles("Title").LanguageID = wdSpanish 
MsgBox ActiveDocument.Styles("Title").Description
```


## See also


#### Concepts


[Style Object](style-object-word.md)

