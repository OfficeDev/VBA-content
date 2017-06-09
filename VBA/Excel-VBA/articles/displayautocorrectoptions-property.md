---
title: DisplayAutoCorrectOptions Property
keywords: vbagr10.chm3086995
f1_keywords:
- vbagr10.chm3086995
ms.prod: excel
api_name:
- Excel.DisplayAutoCorrectOptions
ms.assetid: 9264f123-b3f8-aebc-bfa5-9a3b9be98706
ms.date: 06/08/2017
---


# DisplayAutoCorrectOptions Property

Allows the user to display or hide the AutoCorrect Options button. The default value is True. Read/write Boolean.

 _expression_. **DisplayAutoCorrectOptions**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Remarks

The  **DisplayAutoCorrectOptions** property is a Microsoft Office-wide setting. Changing this property in Microsoft Graph will affect the other Office applications also.


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


