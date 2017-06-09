---
title: Pane.AutoScroll Method (Word)
keywords: vbawd10.chm157286504
f1_keywords:
- vbawd10.chm157286504
ms.prod: word
api_name:
- Word.Pane.AutoScroll
ms.assetid: c0f35128-c98e-2a9e-0ce4-3386c9db89ee
ms.date: 06/08/2017
---


# Pane.AutoScroll Method (Word)

Scrolls automatically through the specified pane.


## Syntax

 _expression_ . **AutoScroll**( **_Velocity_** )

 _expression_ Required. A variable that represents a **[Pane](pane-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Velocity_|Required| **Long**|The speed for scrolling. Can be a number from ? 100 through 100. Use ? 100 for full-speed backward scrolling, and use 100 for full-speed forward scrolling.|

## Remarks

This method continues to run until you stop it manually by pressing a key or clicking the mouse.


## Example

This example scrolls backward through the active window pane slowly.


```vb
ActiveDocument.ActiveWindow.ActivePane.AutoScroll _ 
 Velocity:=-20
```

This example scrolls forward through the active window pane at full speed.




```vb
ActiveDocument.ActiveWindow.ActivePane.AutoScroll _ 
 Velocity:=100
```


## See also


#### Concepts


[Pane Object](pane-object-word.md)

