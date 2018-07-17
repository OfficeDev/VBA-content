---
title: Application.UserInitials Property (Word)
keywords: vbawd10.chm158335029
f1_keywords:
- vbawd10.chm158335029
ms.prod: word
api_name:
- Word.Application.UserInitials
ms.assetid: 00f7d562-4ce5-00e1-bf9d-4325d47947b2
ms.date: 06/08/2017
---


# Application.UserInitials Property (Word)

Returns or sets the user's initials, which Microsoft Word uses to construct comment marks. Read/write  **String** .


## Syntax

 _expression_ . **UserInitials**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Example

This example sets the user's initials.


```vb
Application.UserInitials = "baa"
```

This example returns the letters found in the  **Initials** box on the **User Information** tab in the **Options** dialog box ( **Tools** menu).




```
Msgbox Application.UserInitials
```


## See also


#### Concepts


[Application Object](application-object-word.md)

