---
title: TabControl.Pages Property (Access)
keywords: vbaac10.chm12070
f1_keywords:
- vbaac10.chm12070
ms.prod: access
api_name:
- Access.TabControl.Pages
ms.assetid: dc628cfa-9550-36e6-0aa1-06cf5e80fa25
ms.date: 06/08/2017
---


# TabControl.Pages Property (Access)

Returns a  **[Pages](pages-object-access.md)** collection that represents the pages in the specified **TabControl** object. Read-only.


## Syntax

 _expression_. **Pages**

 _expression_ A variable that represents a **TabControl** object.


## Example

The following example displays a message indicating the number of tabbed pages on tab control TabCtl1.


```vb
MsgBox "Number of pages in TabCtl1:" &; TabCtl1.Pages.Count
```


## See also


#### Concepts


[TabControl Object](tabcontrol-object-access.md)

