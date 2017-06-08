---
title: Application.UserAddress Property (Word)
keywords: vbawd10.chm158335030
f1_keywords:
- vbawd10.chm158335030
ms.prod: word
api_name:
- Word.Application.UserAddress
ms.assetid: 34f9bf48-8af4-4017-a648-13ab7612ca4a
ms.date: 06/08/2017
---


# Application.UserAddress Property (Word)

Returns or sets the user's mailing address. Read/write  **String** .


## Syntax

 _expression_ . **UserAddress**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Remarks

The mailing address is used as a return address on envelopes.


## Example

This example sets the user's return address. The Chr function is used to return a line feed character.


```vb
Application.UserAddress = "4200 Third Street NE" &; Chr(10) _ 
 &; "Anytown, WA 98999"
```

This example returns the address found in the  **Mailing address** box on the **User Information** tab in the **Options** dialog box ( **Tools** menu).




```
Msgbox Application.UserAddress
```


## See also


#### Concepts


[Application Object](application-object-word.md)

