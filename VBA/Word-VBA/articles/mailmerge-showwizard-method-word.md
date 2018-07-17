---
title: MailMerge.ShowWizard Method (Word)
keywords: vbawd10.chm153092210
f1_keywords:
- vbawd10.chm153092210
ms.prod: word
api_name:
- Word.MailMerge.ShowWizard
ms.assetid: 002e6582-4600-c897-f475-546375416cf4
ms.date: 06/08/2017
---


# MailMerge.ShowWizard Method (Word)

Displays the Mail Merge Wizard in a document.


## Syntax

 _expression_ . **ShowWizard**( **_InitialState_** , **_ShowDocumentStep_** , **_ShowTemplateStep_** , **_ShowDataStep_** , **_ShowWriteStep_** , **_ShowPreviewStep_** , **_ShowMergeStep_** )

 _expression_ Required. A variable that represents a **[MailMerge](mailmerge-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _InitialState_|Required| **Variant**|The number of the Mail Merge Wizard step to display.|
| _ShowDocumentStep_|Optional| **Variant**| **True** keeps the "Select document type" step in the sequence of mail merge steps. **False** removes step one.|
| _ShowTemplateStep_|Optional| **Variant**| **True** keeps the "Select starting document" step in the sequence of mail merge steps. **False** removes step two.|
| _ShowDataStep_|Optional| **Variant**| **True** keeps the "Select recipients" step in the sequence of mail merge steps. **False** removes step three.|
| _ShowWriteStep_|Optional| **Variant**| **True** keeps the "Write your letter" step in the sequence of mail merge steps. **False** removes step four.|
| _ShowPreviewStep_|Optional| **Variant**| **True** keeps the "Preview your letters" step in the sequence of mail merge steps. **False** removes step five.|
| _ShowMergeStep_|Optional| **Variant**| **True** keeps the "Complete the merge" step in the sequence of mail merge steps. **False** removes step six.|

## Example

This example checks if the Mail Merge Wizard is already displayed and, if it is, moves to the Mail Merge Wizard's sixth step and removes the fifth step from the Wizard.


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

