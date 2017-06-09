---
title: EmailAuthor.Style Property (Word)
keywords: vbawd10.chm165085287
f1_keywords:
- vbawd10.chm165085287
ms.prod: word
api_name:
- Word.EmailAuthor.Style
ms.assetid: e60dadf7-affd-3bcf-e4a9-d4f083bca000
ms.date: 06/08/2017
---


# EmailAuthor.Style Property (Word)

Returns a  **Style** object that represents the style associated with the current e-mail author for unsent replies, forwards, or new e-mail messages.


## Syntax

 _expression_ . **Style**

 _expression_ Required. A variable that represents an **[EmailAuthor](emailauthor-object-word.md)** object.


## Example

This example returns the style associated with the current author for unsent replies, forwards, or new e-mail messages and displays the name of the font associated with this style.


```vb
Set MyEmailStyle = _ 
 ActiveDocument.Email.CurrentEmailAuthor.Style 
Msgbox MyEmailStyle.Font.Name
```


## See also


#### Concepts


[EmailAuthor Object](emailauthor-object-word.md)

