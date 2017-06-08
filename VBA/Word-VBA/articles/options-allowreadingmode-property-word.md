---
title: Options.AllowReadingMode Property (Word)
keywords: vbawd10.chm162988489
f1_keywords:
- vbawd10.chm162988489
ms.prod: word
api_name:
- Word.Options.AllowReadingMode
ms.assetid: c570b6e8-9d38-7fd5-7cdb-fcd1743bbfe0
ms.date: 06/08/2017
---


# Options.AllowReadingMode Property (Word)

 **True** indicates that Microsoft Word opens documents in Reading Layout view. Read/write **Boolean** .


## Syntax

 _expression_ . **AllowReadingMode**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Remarks

Corresponds to the  **Allow starting in Reading Layout** check box on the **General** tab of the **Options** dialog box.


## Example

The following example toggles the Allow starting in Reading Layout check box.


```vb
Sub ToggleReadingMode() 
 If Options.AllowReadingMode = True Then 
 Options.AllowReadingMode = False 
 Else 
 Options. = True 
 End If 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

