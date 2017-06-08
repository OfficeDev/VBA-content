---
title: Windows.Arrange Method (Word)
keywords: vbawd10.chm157351947
f1_keywords:
- vbawd10.chm157351947
ms.prod: word
api_name:
- Word.Windows.Arrange
ms.assetid: 11325f30-7d28-84f5-4e39-fec34509201e
ms.date: 06/08/2017
---


# Windows.Arrange Method (Word)

Arranges all open document windows in the application workspace.


## Syntax

 _expression_ . **Arrange**( **_ArrangeStyle_** )

 _expression_ A variable that represents a **[Windows](windows-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ArrangeStyle_|Optional| **Variant**|The window arrangement. Can be either of the following  **WdArrangeStyle** constants: **wdIcons** or **wdTiled** .|

## Remarks

Because Microsoft Word uses a Single Document Interface (SDI), this method no longer has any effect.


## Example

This example arranges all open windows so that they don't overlap.


```
Windows.Arrange ArrangeStyle:=wdTiled
```

This example minimizes all open windows and then arranges the minimized windows.




```vb
Dim windowLoop As Window 
 
For Each windowLoop In Windows 
 With windowLoop 
 .Activate 
 .WindowState = wdWindowStateMinimize 
 End With 
Next windowLoop 
 
Windows.Arrange ArrangeStyle:=wdIcons
```


## See also


#### Concepts


[Windows Collection Object](windows-object-word.md)

