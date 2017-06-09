---
title: Window.HorizontalPercentScrolled Property (Word)
keywords: vbawd10.chm157417495
f1_keywords:
- vbawd10.chm157417495
ms.prod: word
api_name:
- Word.Window.HorizontalPercentScrolled
ms.assetid: 18b61708-eb2d-41e0-5b42-9ceb825867e1
ms.date: 06/08/2017
---


# Window.HorizontalPercentScrolled Property (Word)

Returns or sets the horizontal scroll position as a percentage of the document width. Read/write  **Long** .


## Syntax

 _expression_ . **HorizontalPercentScrolled**

 _expression_ A variable that represents a **[Window](window-object-word.md)** object.


## Example

This example displays the percentage that the active window is scrolled horizontally.


```vb
MsgBox _ 
 ActiveDocument.ActiveWindow.HorizontalPercentScrolled &; "%"
```


## See also


#### Concepts


[Window Object](window-object-word.md)

