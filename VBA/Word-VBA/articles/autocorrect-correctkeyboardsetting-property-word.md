---
title: AutoCorrect.CorrectKeyboardSetting Property (Word)
keywords: vbawd10.chm155779090
f1_keywords:
- vbawd10.chm155779090
ms.prod: word
api_name:
- Word.AutoCorrect.CorrectKeyboardSetting
ms.assetid: 2b611e7d-b0fe-41c2-1b93-3364c5d26c9b
ms.date: 06/08/2017
---


# AutoCorrect.CorrectKeyboardSetting Property (Word)

 **True** if Microsoft Word automatically transposes words to their native alphabet if you type text in a language other than the current keyboard language. Read/write **Boolean** .


## Syntax

 _expression_ . **CorrectKeyboardSetting**

 _expression_ An expression that returns an **[AutoCorrect](autocorrect-object-word.md)** object.


## Remarks

The  **[CheckLanguage](application-checklanguage-property-word.md)** property must be set to **True** to use the **CorrectKeyboardSetting** property.


## Example

This example displays a dialog box where the user can choose whether or not Word automatically transposes foreign words to their native alphabets.


```vb
x = MsgBox("Do you want Microsoft Word to tranpose " _ 
 &; "foreign words to their native alphabet?", _ 
 vbYesNo) 
If x = vbYes Then 
 Application.CheckLanguage = True 
 AutoCorrect.CorrectKeyboardSetting = True 
 MsgBox "Automatic keyboard correction enabled!" 
End If
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-word.md)

