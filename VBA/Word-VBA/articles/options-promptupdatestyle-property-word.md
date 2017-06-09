---
title: Options.PromptUpdateStyle Property (Word)
keywords: vbawd10.chm162988473
f1_keywords:
- vbawd10.chm162988473
ms.prod: word
api_name:
- Word.Options.PromptUpdateStyle
ms.assetid: 0646e8e2-3462-14c7-7e73-d35ad9a20724
ms.date: 06/08/2017
---


# Options.PromptUpdateStyle Property (Word)

 **True** displays a message asking the user to verify whether they want to reformat a style or reapply the original style formatting when changing the formatting of styles. Read/write **Boolean** .


## Syntax

 _expression_ . **PromptUpdateStyle**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Remarks

 **False** reapplies the style formatting to the selection without verifying whether the user wants to change the style.


## Example

This example checks to see if a user receives a message when updating styles, and if not, enables it.


```vb
Sub UpdateStylePrompt() 
 With Application.Options 
 If .PromptUpdateStyle = False Then 
 .PromptUpdateStyle = True 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

