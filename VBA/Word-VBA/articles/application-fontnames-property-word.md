---
title: Application.FontNames Property (Word)
keywords: vbawd10.chm158334987
f1_keywords:
- vbawd10.chm158334987
ms.prod: word
api_name:
- Word.Application.FontNames
ms.assetid: 6aeadf51-79c7-1123-ea64-582ceee26443
ms.date: 06/08/2017
---


# Application.FontNames Property (Word)

Returns a  **[FontNames](fontnames-object-word.md)** object that includes the names of all the available fonts. Read-only.


## Syntax

 _expression_ . **FontNames**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example displays the font names in the  **FontNames** collection.


```vb
Dim strFont As String 
Dim intResponse As Integer 
 
For Each strFont In FontNames 
 intResponse = MsgBox(Prompt:=strFont, Buttons:=vbOKCancel) 
 If intResponse = vbCancel Then Exit For 
Next strFont
```


## See also


#### Concepts


[Application Object](application-object-word.md)

