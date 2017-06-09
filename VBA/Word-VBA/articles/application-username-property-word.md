---
title: Application.UserName Property (Word)
keywords: vbawd10.chm158335028
f1_keywords:
- vbawd10.chm158335028
ms.prod: word
api_name:
- Word.Application.UserName
ms.assetid: 96f5ffb6-a20d-96f0-e3a4-0ad2dd47bf99
ms.date: 06/08/2017
---


# Application.UserName Property (Word)

Returns or sets the user's name, which is used on envelopes and for the Author document property. Read/write  **String** .


## Syntax

 _expression_ . **UserName**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Example

This example sets the user's name.


```vb
Application.UserName = "Andrew Fuller"
```

This example returns the name found in the  **Name** box on the **User Information** tab in the **Options** dialog box ( **Tools** menu).




```
Msgbox Application.UserName
```


## See also


#### Concepts


[Application Object](application-object-word.md)

