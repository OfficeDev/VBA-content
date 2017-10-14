---
title: EmailOptions.MarkComments Property (Word)
keywords: vbawd10.chm165347435
f1_keywords:
- vbawd10.chm165347435
ms.prod: word
api_name:
- Word.EmailOptions.MarkComments
ms.assetid: 792e77b2-ba00-2b2b-c81b-7d00dad702cd
ms.date: 06/08/2017
---


# EmailOptions.MarkComments Property (Word)

 **True** if Microsoft Word marks the user's comments in e-mail messages. Read/write **Boolean** .


## Syntax

 _expression_ . **MarkComments**

 _expression_ An expression that returns an **[EmailOptions](emailoptions-object-word.md)** object.


## Remarks

This property marks comments with the value of the  **[MarkCommentsWith](emailoptions-markcommentswith-property-word.md)** property. The default value of the **MarkCommentsWith** property is the value of the **[UserName](application-username-property-word.md)** property.


## Example

This example sets Word to mark comments in e-mail messages with the initials "WK."


```vb
Application.EmailOptions.MarkCommentsWith = "WK" 
Application.EmailOptions.MarkComments = True
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

