---
title: Options.CheckHangulEndings Property (Word)
keywords: vbawd10.chm162988373
f1_keywords:
- vbawd10.chm162988373
ms.prod: word
api_name:
- Word.Options.CheckHangulEndings
ms.assetid: fdb1e463-62d9-7053-13b2-e5dec345912e
ms.date: 06/08/2017
---


# Options.CheckHangulEndings Property (Word)

 **True** if Microsoft Word automatically detects Hangul endings and ignores them during conversion from Hangul to Hanja. Read/write **Boolean** .


## Syntax

 _expression_ . **CheckHangulEndings**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Remarks

If converting from Hanja to Hangul, this property is ignored.


## Example

This example asks the user whether to set Microsoft Word to automatically detect Hangul endings and ignore them during conversion from Hangul to hanja.


```vb
x = MsgBox("Check Hangul endings during " _ 
 &; "conversion from Hangul to Hanja?", vbYesNo) 
If x = vbYes Then 
 Options.CheckHangulEndings = True 
Else 
 Options.CheckHangulEndings = False 
End If
```


## See also


#### Concepts


[Options Object](options-object-word.md)

