---
title: ProtectedViewWindow.EnableResize Property (Excel)
keywords: vbaxl10.chm914075
f1_keywords:
- vbaxl10.chm914075
ms.prod: excel
api_name:
- Excel.ProtectedViewWindow.EnableResize
ms.assetid: 110c4080-7dea-e34d-224b-47337e5d6777
ms.date: 06/08/2017
---


# ProtectedViewWindow.EnableResize Property (Excel)

 **True** if the **Protected View** window can be resized. Read/write


## Syntax

 _expression_ . **EnableResize**

 _expression_ A variable that represents a **[ProtectedViewWindow](protectedviewwindow-object-excel.md)** object.


### Return Value

 **Boolean**


## Example

The following code example sets the active  **Protected View** window so that it cannot be resized.


```vb
ActiveProtectedViewWindow.EnableResize = False
```


## See also


#### Concepts


[ProtectedViewWindow Object](protectedviewwindow-object-excel.md)

