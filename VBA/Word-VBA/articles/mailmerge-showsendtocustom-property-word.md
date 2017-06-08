---
title: MailMerge.ShowSendToCustom Property (Word)
keywords: vbawd10.chm153092109
f1_keywords:
- vbawd10.chm153092109
ms.prod: word
api_name:
- Word.MailMerge.ShowSendToCustom
ms.assetid: 261d5edc-8320-7f73-0b78-899898834c35
ms.date: 06/08/2017
---


# MailMerge.ShowSendToCustom Property (Word)

Returns or sets a  **String** corresponding to the caption on a custom button on the Complete the merge step (step six) of the Mail Merge Wizard. Read/write.


## Syntax

 _expression_ . **ShowSendToCustom**

 _expression_ A variable that represents a **[MailMerge](mailmerge-object-word.md)** object.


## Remarks

When a user clicks the custom button, the  **[MailMergeWizardSendToCustom](application-mailmergewizardsendtocustom-event-word.md)** event executes.


## Example

This example displays a custom button on the sixth step of the Mail Merge Wizard only for mailing labels.


```vb
Sub ShowCustomButton() 
 With ActiveDocument.MailMerge 
 If .MainDocumentType = wdMailingLabels Then 
 .ShowSendToCustom = "Custom Label Processing" 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[MailMerge Object](mailmerge-object-word.md)

