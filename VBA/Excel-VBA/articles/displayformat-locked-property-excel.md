---
title: DisplayFormat.Locked Property (Excel)
keywords: vbaxl10.chm893082
f1_keywords:
- vbaxl10.chm893082
ms.prod: excel
api_name:
- Excel.DisplayFormat.Locked
ms.assetid: 32941867-c714-cfa1-ad16-c214e745580e
ms.date: 06/08/2017
---


# DisplayFormat.Locked Property (Excel)

Returns a value that indicates if the associated  **[Range](range-object-excel.md)** object is locked as it is displayed in the current user interface. Read-only.


## Syntax

 _expression_ . **Locked**

 _expression_ A variable that represents a **[DisplayFormat](displayformat-object-excel.md)** object.


### Return Value

Variant


## Remarks

Returns  **True** if the range is locked, **False** if the range can be modified when the sheet is protected, or **Null** if the range contains both locked and unlocked cells.


## See also


#### Concepts


[DisplayFormat Object](displayformat-object-excel.md)

