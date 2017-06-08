---
title: Pane.NewFrameset Method (Word)
keywords: vbawd10.chm157286506
f1_keywords:
- vbawd10.chm157286506
ms.prod: word
api_name:
- Word.Pane.NewFrameset
ms.assetid: 86724851-6b29-1a66-e863-edeb4c9d43de
ms.date: 06/08/2017
---


# Pane.NewFrameset Method (Word)

Creates a new frames page based on the specified pane.


## Syntax

 _expression_ . **NewFrameset**

 _expression_ Required. A variable that represents a **[Pane](pane-object-word.md)** object.


## Remarks

For more information on creating frames pages, see [Creating frames pages](http://msdn.microsoft.com/library/0245564e-b2df-83cd-1e32-e63079970dc1%28Office.15%29.aspx).


## Example

This example opens a document named "Temp.doc" and then creates a new frames page whose only frame contains "Temp.doc".


```
Documents.Open "C:\Documents\Temp.doc" 
ActiveDocument.ActiveWindow.ActivePane.NewFrameset
```


## See also


#### Concepts


[Pane Object](pane-object-word.md)

