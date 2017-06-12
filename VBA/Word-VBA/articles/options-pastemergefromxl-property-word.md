---
title: Options.PasteMergeFromXL Property (Word)
keywords: vbawd10.chm162988466
f1_keywords:
- vbawd10.chm162988466
ms.prod: word
api_name:
- Word.Options.PasteMergeFromXL
ms.assetid: d09c2244-71f5-3345-fcbe-14a307f23da3
ms.date: 06/08/2017
---


# Options.PasteMergeFromXL Property (Word)

 **True** to merge table formatting when pasting from Microsoft Excel. Read/write **Boolean** .


## Syntax

 _expression_ . **PasteMergeFromXL**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Example

This example sets Microsoft Word to automatically merge table formatting when pasting Excel tables if the option has been disabled.


```vb
Sub AdjustExcelFormatting() 
 With Options 
 If .PasteMergeFromXL = False Then 
 .PasteMergeFromXL = True 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

