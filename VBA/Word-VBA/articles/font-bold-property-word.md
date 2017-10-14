---
title: Font.Bold Property (Word)
keywords: vbawd10.chm156369026
f1_keywords:
- vbawd10.chm156369026
ms.prod: word
api_name:
- Word.Font.Bold
ms.assetid: 84e8d6b7-1be4-e1c5-c246-a6370011bc8b
ms.date: 06/08/2017
---


# Font.Bold Property (Word)

 **True** if the font is formatted as bold. Read/write **Long** .


## Syntax

 _expression_ . **Bold**

 _expression_ A variable that represents a **[Font](font-object-word.md)** object.


## Remarks

Returns  **True** , **False** or **wdUndefined** (a mixture of **True** and **False** ). Can be set to **True** , **False** , or **wdToggle** .


## Example

This example makes the entire selection bold if part of the selection is formatted as bold.


```vb
If Selection.Type = wdSelectionNormal Then 
 If Selection.Font.Bold = wdUndefined Then _ 
 Selection.Font.Bold = True 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also


#### Concepts


[Font Object](font-object-word.md)

