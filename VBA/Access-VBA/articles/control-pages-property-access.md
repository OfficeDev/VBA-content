---
title: Control.Pages Property (Access)
keywords: vbaac10.chm10149
f1_keywords:
- vbaac10.chm10149
ms.prod: access
api_name:
- Access.Control.Pages
ms.assetid: fd4ea2c0-ea8c-51a0-a012-8ba5848d3516
ms.date: 06/08/2017
---


# Control.Pages Property (Access)

Returns a  **[Pages](pages-object-access.md)** collection that represents the pages in the specified control that supports tabbed pages (for example, a **TabControl** object). Read-only.


## Syntax

 _expression_. **Pages**

 _expression_ A variable that represents a **Control** object.


## Example

The following example displays a message indicating the number of tabbed pages on tab control TabCtl1.


```vb
MsgBox "Number of pages in TabCtl1:" &; TabCtl1.Pages.Count
```


## See also


#### Concepts


[Control Object](control-object-access.md)

