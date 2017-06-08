---
title: MailMerge.DocumentUpdating Property (Publisher)
keywords: vbapb10.chm6225925
f1_keywords:
- vbapb10.chm6225925
ms.prod: publisher
api_name:
- Publisher.MailMerge.DocumentUpdating
ms.assetid: c65ca4a0-e5eb-d97e-9126-4af86f4e805f
ms.date: 06/08/2017
---


# MailMerge.DocumentUpdating Property (Publisher)

Returns or sets a  **Boolean** indicating whether the screen is updated while mail merge code is running. Default is **True** (the screen is updated). Read/write.


## Syntax

 _expression_. **DocumentUpdating**

 _expression_A variable that represents a  **MailMerge** object.


### Return Value

Boolean


## Remarks

Turning document updating off during run time can speed execution of Microsoft Visual Basic code. However, we recommend that you provide some indication of status so that the user is aware that the program is functioning correctly.


## Example

The following example turns off document updating at the beginning of a mail merge subroutine and turns it back on at the end of the subroutine.


```vb
Sub MailMergeProcedure() 
 ActiveDocument.MailMerge.DocumentUpdating = False ' Mail merge code. 
ActiveDocument.MailMerge.DocumentUpdating = True 
End Sub
```


