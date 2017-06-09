---
title: Chart.DisplayBlanksAs Property (Word)
keywords: vbawd10.chm79364114
f1_keywords:
- vbawd10.chm79364114
ms.prod: word
api_name:
- Word.Chart.DisplayBlanksAs
ms.assetid: 573752ec-7c2a-a5e0-bd05-626c81fb5d48
ms.date: 06/08/2017
---


# Chart.DisplayBlanksAs Property (Word)

Returns or sets the way that blank cells are plotted on a chart. Can be one of the  **[XlDisplayBlanksAs](xldisplayblanksas-enumeration-word.md)** constants. Read/write **Long** .


## Syntax

 _expression_ . **DisplayBlanksAs**

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


## Example

The following example sets Microsoft Word to not plot blank cells for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.DisplayBlanksAs = xlNotPlotted 
 End If 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

