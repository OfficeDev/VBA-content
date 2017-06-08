---
title: Options.EnableSound Property (Word)
keywords: vbawd10.chm162988053
f1_keywords:
- vbawd10.chm162988053
ms.prod: word
api_name:
- Word.Options.EnableSound
ms.assetid: c7934437-2d32-2a2a-9eab-c0dac74b2108
ms.date: 06/08/2017
---


# Options.EnableSound Property (Word)

 **True** if Word makes the computer respond with a sound whenever an error occurs. Read/write **Boolean** .


## Syntax

 _expression_ . **EnableSound**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Example

This example sets the  **Provide feedback with sound** option on the **General** tab in the **Options** dialog box, based on user input.


```vb
If MsgBox("Do you want Word to beep on errors?", 36) = vbYes Then 
 Options.EnableSound = True 
Else 
 Options.EnableSound = False 
End If
```


## See also


#### Concepts


[Options Object](options-object-word.md)

