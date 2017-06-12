---
title: CheckBox.Default Property (Word)
keywords: vbawd10.chm153485315
f1_keywords:
- vbawd10.chm153485315
ms.prod: word
api_name:
- Word.CheckBox.Default
ms.assetid: 49e27047-aee0-bf84-ce44-7d30d7f863e8
ms.date: 06/08/2017
---


# CheckBox.Default Property (Word)

Returns or sets the default check box value.  **True** if the default value is checked. Read/write **Boolean** .


## Syntax

 _expression_ . **Default**

 _expression_ Required. A variable that represents a **[CheckBox](checkbox-object-word.md)** object.


## Example

If the first form field in the active document is a check box, this example retrieves the default value.


```vb
Dim blnDefault As Boolean 
 
If ActiveDocument.FormFields(1).Type = wdFieldFormCheckBox Then 
 blnDefault = ActiveDocument.FormFields(1).CheckBox.DefaultEnd If
```


## See also


#### Concepts


[CheckBox Object](checkbox-object-word.md)

