---
title: WebOptions.ScreenSize Property (Word)
keywords: vbawd10.chm165937160
f1_keywords:
- vbawd10.chm165937160
ms.prod: word
api_name:
- Word.WebOptions.ScreenSize
ms.assetid: 4398a153-6932-17ef-b449-a532363fb428
ms.date: 06/08/2017
---


# WebOptions.ScreenSize Property (Word)

Returns or sets the ideal minimum screen size (width by height, in pixels) that you should use when viewing the saved document in a Web browser. Read/write  **MsoScreenSize** .


## Syntax

 _expression_ . **ScreenSize**

 _expression_ Required. A variable that represents a **[WebOptions](weboptions-object-word.md)** collection.


## Example

This example sets the target screen size for the active Web page at 800x600 pixels.


```vb
ActiveDocument.WebOptions.ScreenSize = _ 
 msoScreenSize800x600
```


## See also


#### Concepts


[WebOptions Object](weboptions-object-word.md)

