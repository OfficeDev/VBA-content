---
title: Slicer.DisableMoveResizeUI Property (Excel)
keywords: vbaxl10.chm905077
f1_keywords:
- vbaxl10.chm905077
ms.prod: excel
api_name:
- Excel.Slicer.DisableMoveResizeUI
ms.assetid: 2477e495-e61a-6981-6df2-5bb1cb480576
ms.date: 06/08/2017
---


# Slicer.DisableMoveResizeUI Property (Excel)

Returns or sets whether the specified slicer can be moved or resized by using the user interface. Read/write.


## Syntax

 _expression_ . **DisableMoveResizeUI**

 _expression_ A variable that represents a **[Slicer](slicer-object-excel.md)** object.


### Return Value

Boolean


## Remarks

 **True** if the slicer cannot be moved or resized by selecting borders or handles in the user interface; otherwise **False** . The default value is **False** . Setting the **DisableMoveResizeUI** property to **True** affects only the user interface. Moving or resizing the slicer by setting properties such as the **[Top](slicer-top-property-excel.md)** , **[Left](slicer-left-property-excel.md)** , **[Width](slicer-width-property-excel.md)** , or **[Height](slicer-height-property-excel.md)** properties of the **Slicer** object from code is not disabled.


## See also


#### Concepts


[Slicer Object](slicer-object-excel.md)

