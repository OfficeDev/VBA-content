---
title: Panes.Add Method (Word)
keywords: vbawd10.chm157220867
f1_keywords:
- vbawd10.chm157220867
ms.prod: word
api_name:
- Word.Panes.Add
ms.assetid: 34dba7e0-cb4f-0482-c8c5-cc3d54cacc9c
ms.date: 06/08/2017
---


# Panes.Add Method (Word)

Returns a  **Pane** object that represents a new pane to a window.


## Syntax

 _expression_ . **Add**( **_SplitVertical_** )

 _expression_ Required. A variable that represents a **[Panes](panes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SplitVertical_|Optional| **Variant**|A number that represents the percentage of the window, from top to bottom, you want to appear above the split.|

### Return Value

Pane


## Remarks

This method will fail if it is applied to a window that has already been split.


## Example

The following example splits the active window such that the top pane is 30 percent of the total window size.


```vb
ActiveDocument.ActiveWindow.Panes.Add SplitVertical:=30
```


## See also


#### Concepts


[Panes Collection Object](panes-object-word.md)

