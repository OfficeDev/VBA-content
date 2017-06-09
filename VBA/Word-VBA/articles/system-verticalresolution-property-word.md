---
title: System.VerticalResolution Property (Word)
keywords: vbawd10.chm154468360
f1_keywords:
- vbawd10.chm154468360
ms.prod: word
api_name:
- Word.System.VerticalResolution
ms.assetid: f93b0eed-1b0c-654c-8c73-60da0d13ab11
ms.date: 06/08/2017
---


# System.VerticalResolution Property (Word)

Returns the vertical screen resolution in pixels. Read-only  **Long** .


## Syntax

 _expression_ . **VerticalResolution**

 _expression_ An expression that returns a **[System](system-object-word.md)** object.


## Example

This example displays the current screen resolution (for example, "1024 x 768").


```
horz = System.HorizontalResolution 
vert = System.VerticalResolution 
MsgBox "Resolution = " &; horz &; " x " &; vert
```


## See also


#### Concepts


[System Object](system-object-word.md)

