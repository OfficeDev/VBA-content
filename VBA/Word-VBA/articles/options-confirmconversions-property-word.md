---
title: Options.ConfirmConversions Property (Word)
keywords: vbawd10.chm162988054
f1_keywords:
- vbawd10.chm162988054
ms.prod: word
api_name:
- Word.Options.ConfirmConversions
ms.assetid: 4bdea504-e8c7-193c-c671-1a8ec84d93ca
ms.date: 06/08/2017
---


# Options.ConfirmConversions Property (Word)

 **True** if Word displays the **Convert File** dialog box before it opens or inserts a file that isn't a Word document or template. In the **Convert File** dialog box, the user chooses the format to convert the file from. Read/write **Boolean** .


## Syntax

 _expression_ . **ConfirmConversions**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Example

This example sets Word to display the  **Convert File** dialog box whenever a file that isn't a Word document or template is opened.


```vb
Options.ConfirmConversions = True
```

This example returns the current status of the  **Confirm conversion at Open** option on the **General** tab in the **Options** dialog box.




```vb
Dim blnConfirm As Boolean 
 
blnConfirm= Options.ConfirmConversions
```


## See also


#### Concepts


[Options Object](options-object-word.md)

