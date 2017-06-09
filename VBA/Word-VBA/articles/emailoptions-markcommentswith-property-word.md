---
title: EmailOptions.MarkCommentsWith Property (Word)
keywords: vbawd10.chm165347434
f1_keywords:
- vbawd10.chm165347434
ms.prod: word
api_name:
- Word.EmailOptions.MarkCommentsWith
ms.assetid: f10ce322-5ac5-f431-80c9-5c00a0892e2e
ms.date: 06/08/2017
---


# EmailOptions.MarkCommentsWith Property (Word)

Returns or sets the string with which Microsoft Word marks comments in e-mail messages. Read/write  **String** .


## Syntax

 _expression_ . **MarkCommentsWith**

 _expression_ An expression that returns an **[EmailOptions](emailoptions-object-word.md)** object.


## Remarks

The default value is the value of the  **[UserName](application-username-property-word.md)** property.


## Example

This example sets Word to mark comments in e-mail messages with the initials "WK."


```vb
Application.EmailOptions.MarkCommentsWith = "WK" 
Application.EmailOptions.MarkComments = True
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

