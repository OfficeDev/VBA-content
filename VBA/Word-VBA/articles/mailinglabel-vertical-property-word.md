---
title: MailingLabel.Vertical Property (Word)
keywords: vbawd10.chm152502282
f1_keywords:
- vbawd10.chm152502282
ms.prod: WORD
api_name:
- Word.MailingLabel.Vertical
ms.assetid: 9dac957c-d2be-addd-81f2-4dd6b134051d
---


# MailingLabel.Vertical Property (Word)

 **True** vertically orients text on Asian mailing labels. Read/write **Boolean** .


## Syntax

 _expression_ . **Vertical**

 _expression_ Required. A variable that represents a **[MailingLabel](mailinglabel-object-word.md)** object.


## Example

This example determines if the active document is a mail merge mailing label document and if the language setting is Japanese, and if so, sets the mailing label's orientation to vertical.


```vb
Sub VerticalLabel() 
 If ActiveDocument.MailMerge.MainDocumentType = wdMailingLabels And 
 Application.Language = msoLanguageIDJapanese Then 
 Application.MailingLabel.Vertical = True 
 End If 
End Sub
```


## See also


#### Concepts


[MailingLabel Object](mailinglabel-object-word.md)

