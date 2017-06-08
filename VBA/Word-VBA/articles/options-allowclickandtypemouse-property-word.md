---
title: Options.AllowClickAndTypeMouse Property (Word)
keywords: vbawd10.chm162988445
f1_keywords:
- vbawd10.chm162988445
ms.prod: word
api_name:
- Word.Options.AllowClickAndTypeMouse
ms.assetid: 40b6f33c-a577-ff1e-6f7c-46b971e34cab
ms.date: 06/08/2017
---


# Options.AllowClickAndTypeMouse Property (Word)

 **True** if Click and Type functionality is enabled. Read/write **Boolean** .


## Syntax

 _expression_ . **AllowClickAndTypeMouse**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Remarks

For more information on Click and Type, see About Click and Type .


## Example

This example checks to determine whether Click and Type functionality is enabled. If it isn't enabled, the example sets this functionality based on the user's choice.


```vb
If Options.AllowClickAndTypeMouse = False Then 
 x = MsgBox("Do you want to use Click and Type?", _ 
 vbYesNo) 
 If x = vbYes Then 
 Options.AllowClickAndTypeMouse = True 
 MsgBox "Click and Type enabled!" 
 End If 
End If
```


## See also


#### Concepts


[Options Object](options-object-word.md)

