---
title: CustomLabel.DotMatrix Property (Word)
keywords: vbawd10.chm152371211
f1_keywords:
- vbawd10.chm152371211
ms.prod: word
api_name:
- Word.CustomLabel.DotMatrix
ms.assetid: 46646fd9-2d37-ed2b-d6f9-68cf139bbd57
ms.date: 06/08/2017
---


# CustomLabel.DotMatrix Property (Word)

 **True** if the printer type for the specified custom label is dot matrix. **False** if the printer type is either laser or ink jet. Read-only **Boolean** .


## Syntax

 _expression_ . **DotMatrix**

 _expression_ A variable that represents a **[CustomLabel](customlabel-object-word.md)** object.


## Example

This example displays the name and printer type of the first custom mailing label.


```vb
Dim mlTemp As MailingLabel 
 
Set mlTemp = Application.MailingLabel 
If mlTemp.CustomLabels.Count >= 1 Then 
 If mlTemp.CustomLabels(1).DotMatrix = True Then 
 MsgBox mlTemp.CustomLabels(1).Name &; " is dot matrix" 
 Else 
 MsgBox mlTemp.CustomLabels(1).Name _ 
 &; " is laser or ink jet" 
 End If 
End If
```


## See also


#### Concepts


[CustomLabel Object](customlabel-object-word.md)

