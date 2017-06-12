---
title: Windows.SyncScrollingSideBySide Property (Word)
keywords: vbawd10.chm157352939
f1_keywords:
- vbawd10.chm157352939
ms.prod: word
api_name:
- Word.SyncScrollingSideBySide
ms.assetid: d6d84719-fc49-acd4-acfe-154d2b45b74a
ms.date: 06/08/2017
---


# Windows.SyncScrollingSideBySide Property (Word)

 **True** enables scrolling the contents of the windows at the same time. Read/write **Boolean** .


## Syntax

 _expression_ . **SyncScrollingSideBySide**

 _expression_ An expression that returns a **Windows** collection.


## Remarks

 **False** disables scrolling the windows at the same time.


## Example

The following example enables scrolling of adjacent windows at the same time.


```vb
Dim objDoc1 As Word.Document 
Dim objDoc2 As Word.Document 
 
Set objDoc1 = Documents.Add 
Set objDoc2 = Documents.Add 
 
objDoc2.Activate 
objDoc2.Windows.CompareSideBySideWith objDoc1 
Windows.ResetPositionsSideBySide 
Windows.SyncScrollingSideBySide = True
```


## See also


#### Concepts


[Windows Collection Object](windows-object-word.md)

