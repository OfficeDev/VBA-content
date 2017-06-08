---
title: MailMerge.WizardState Property (Word)
keywords: vbawd10.chm153092110
f1_keywords:
- vbawd10.chm153092110
ms.prod: word
api_name:
- Word.MailMerge.WizardState
ms.assetid: 7dc71e03-cdc4-c307-d433-1d3984aa39d4
ms.date: 06/08/2017
---


# MailMerge.WizardState Property (Word)

Returns or sets a  **Long** indicating the current Mail Merge Wizard step for a document. The WizardState method returns a number that equates to the current Mail Merge Wizard step; a zero (0) means the Mail Merge Wizard is closed. Read/write.


## Syntax

 _expression_ . **WizardState**

 _expression_ A variable that represents a **[MailMerge](mailmerge-object-word.md)** object.


## Example

This example checks if the Mail Merge Wizard is already displayed in the active document and if it is, moves to the Mail Merge Wizard's sixth step and removes the fifth step from the Wizard.


```vb
Sub ShowMergeWizard() 
 With ActiveDocument.MailMerge 
 If .WizardState > 0 Then 
 .ShowWizard InitialState:=6, ShowPreviewStep:=False 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[MailMerge Object](mailmerge-object-word.md)

