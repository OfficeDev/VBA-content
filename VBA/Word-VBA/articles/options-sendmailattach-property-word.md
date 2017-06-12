---
title: Options.SendMailAttach Property (Word)
keywords: vbawd10.chm162988056
f1_keywords:
- vbawd10.chm162988056
ms.prod: word
api_name:
- Word.Options.SendMailAttach
ms.assetid: e749ca30-089f-5116-ce70-a3d760006a2c
ms.date: 06/08/2017
---


# Options.SendMailAttach Property (Word)

 **True** if the **Send To** command on the **File** menu inserts the active document as an attachment to a mail message. Read/write **Boolean** .


## Syntax

 _expression_ . **SendMailAttach**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Remarks

 **False** if the **Send To** command inserts the contents of the active document as text in a mail message.


## Example

This example opens a new mail message that has the active document attached to it.


```vb
Options.SendMailAttach = True 
ActiveDocument.SendMail
```

This example returns the state of the  **Mail as attachment** option on the **General** tab of the **Options** dialog box.




```
Msgbox Options.SendMailAttach
```


## See also


#### Concepts


[Options Object](options-object-word.md)

