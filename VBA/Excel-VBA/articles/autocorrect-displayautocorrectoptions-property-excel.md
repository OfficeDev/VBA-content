---
title: AutoCorrect.DisplayAutoCorrectOptions Property (Excel)
keywords: vbaxl10.chm545081
f1_keywords:
- vbaxl10.chm545081
ms.prod: excel
api_name:
- Excel.AutoCorrect.DisplayAutoCorrectOptions
ms.assetid: 3f37f82b-468c-9bf7-2bae-4c081a41a888
ms.date: 06/08/2017
---


# AutoCorrect.DisplayAutoCorrectOptions Property (Excel)

Allows the user to display or hide the  **AutoCorrect Options** button. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayAutoCorrectOptions**

 _expression_ A variable that represents an **AutoCorrect** object.


## Remarks

The  **DisplayAutoCorrectOptions** property is a Microsoft Office-wide setting. Changing this property in Microsoft Excel will affect the other Office applications also.

In Excel the  **AutoCorrect Options** button only appears when a hyperlink is automatically created.


## Example

This example determines if the  **AutoCorrect Options** button can be displayed and notifies the user.


```vb
Sub CheckDisplaySetting() 
 
 'Determine setting and notify user. 
 If Application.AutoCorrect.DisplayAutoCorrectOptions = True Then 
 MsgBox "The AutoCorrect Options button can be displayed." 
 Else 
 MsgBox "The AutoCorrect Options button cannot be displayed." 
 End If 
 
End Sub
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-excel.md)

