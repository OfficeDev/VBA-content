---
title: Window.ScrollIntoView Method (Word)
keywords: vbawd10.chm157417583
f1_keywords:
- vbawd10.chm157417583
ms.prod: word
api_name:
- Word.Window.ScrollIntoView
ms.assetid: b16afab5-8645-dfd6-2b4b-8924fe49916a
ms.date: 06/08/2017
---


# Window.ScrollIntoView Method (Word)

Scrolls through the document window so the specified range or shape is displayed in the document window.


## Syntax

 _expression_ . **ScrollIntoView**( **_Obj_** , **_Start_** )

 _expression_ Required. A variable that represents a **[Window](window-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Obj_|Required| **Object**|A  **Range** or **Shape** object.|
| _Start_|Optional| **Boolean**| **True** if the upper-left corner of the range or shape appears at the upper-left corner of the document window. **False** if the lower-right corner of the range or shape appears at the lower-right corner of the document window. The default value is **True** .|

## Remarks

If the range or shape is larger than the document window, the Start argument specifies which portion of the range or shape displays or gets initial focus. This method cannot be used with outline view.


## Example

This example scrolls through the active document so that the current selection is visible in the document window.


```vb
ActiveWindow.ScrollIntoView Selection.Range, True
```


## See also


#### Concepts


[Window Object](window-object-word.md)

