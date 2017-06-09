---
title: MailMerge.WizardState Property (Publisher)
keywords: vbapb10.chm6225929
f1_keywords:
- vbapb10.chm6225929
ms.prod: publisher
api_name:
- Publisher.MailMerge.WizardState
ms.assetid: a237cb3f-2c03-5f62-fa67-d4aa7703389d
ms.date: 06/08/2017
---


# MailMerge.WizardState Property (Publisher)

Returns or sets a  **Long** indicating the current Mail Merge wizard step for a publication. The **WizardState** property returns a number that equates to the current Mail Merge wizard step; a zero (0) means the Mail Merge wizard is closed. Read/write.


## Syntax

 _expression_. **WizardState**

 _expression_A variable that represents a  **MailMerge** object.


### Return Value

Long


## Example

This example displays the Mail Merge wizard if it is closed.


```vb
Sub ShowMergeWizard() 
 With ActiveDocument.MailMerge 
 If .WizardState = 0 Then 
 .ShowWizard 
 End If 
 End With 
End Sub
```


