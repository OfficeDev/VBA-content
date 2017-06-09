---
title: DefaultWebOptions.ScreenSize Property (Word)
keywords: vbawd10.chm165871627
f1_keywords:
- vbawd10.chm165871627
ms.prod: word
api_name:
- Word.DefaultWebOptions.ScreenSize
ms.assetid: 21f1019f-6658-0da9-519e-adefc8356607
ms.date: 06/08/2017
---


# DefaultWebOptions.ScreenSize Property (Word)

Returns or sets the ideal minimum screen size (width by height, in pixels) that you should use when viewing the saved document in a Web browser. Read/write  **MsoScreenSize** .


## Syntax

 _expression_ . **ScreenSize**

 _expression_ Required. A variable that represents a **[DefaultWebOptions](defaultweboptions-object-word.md)** collection.


## Example

This example sets the target screen size at 800x600 pixels.


```vb
Application.DefaultWebOptions.ScreenSize = _ 
 msoScreenSize800x600
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-word.md)

