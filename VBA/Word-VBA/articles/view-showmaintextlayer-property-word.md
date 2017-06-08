---
title: View.ShowMainTextLayer Property (Word)
keywords: vbawd10.chm161808411
f1_keywords:
- vbawd10.chm161808411
ms.prod: word
api_name:
- Word.View.ShowMainTextLayer
ms.assetid: 0e2b3dd8-8e42-5f53-abc0-849daa5683bc
ms.date: 06/08/2017
---


# View.ShowMainTextLayer Property (Word)

 **True** if the text in the specified document is visible when the header and footer areas are displayed. This property is equivalent to the **Show/Hide Document Text** button on the **Header and Footer** toolbar. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowMainTextLayer**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Example

This example displays the document header in the active window and hides the document text.


```vb
With ActiveDocument.ActiveWindow.View 
 .Type = wdPrintView 
 .SeekView = wdSeekCurrentPageHeader 
 .ShowMainTextLayer = False 
End With
```


## See also


#### Concepts


[View Object](view-object-word.md)

