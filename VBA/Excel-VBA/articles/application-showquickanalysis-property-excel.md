---
title: Application.ShowQuickAnalysis Property (Excel)
keywords: vbaxl10.chm133337
f1_keywords:
- vbaxl10.chm133337
ms.prod: excel
ms.assetid: 043d9523-1fbc-0afd-2adf-9775e71058c0
ms.date: 06/08/2017
---


# Application.ShowQuickAnalysis Property (Excel)

Controls whether the Quick Analysis contextual user interface is displayed on selection.  **TRUE** means the Quick Analysis button will show. Corresponds to the **Show Quick Analysis options on selection** checkbox located in the **File** menu, **Options**,  **Excel Options**, and then  **General** tab. Read/Write. **Boolean** .


## Syntax

 _expression_ . **ShowQuickAnalysis**

 _expression_ A variable that represents an **Application** object.


## Example

This example hides the Quick Analysis button on selection.


```vb
Application.ShowQuickAnalysis = False
```


## Property value

 **BOOL**


## See also


#### Concepts


[Application Object](application-object-excel.md)

