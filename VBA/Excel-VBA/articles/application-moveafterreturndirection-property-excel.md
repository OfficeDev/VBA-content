---
title: Application.MoveAfterReturnDirection Property (Excel)
keywords: vbaxl10.chm133169
f1_keywords:
- vbaxl10.chm133169
ms.prod: excel
api_name:
- Excel.Application.MoveAfterReturnDirection
ms.assetid: c11d8e36-755e-c911-de44-8b630b549418
ms.date: 06/08/2017
---


# Application.MoveAfterReturnDirection Property (Excel)

Returns or sets the direction in which the active cell is moved when the user presses ENTER. Read/write  **[XlDirection](xldirection-enumeration-excel.md)** .


## Syntax

 _expression_ . **MoveAfterReturnDirection**

 _expression_ A variable that represents an **Application** object.


## Remarks



| **XlDirection** can be one of these **XlDirection** constants.|
| **xlDown**|
| **xlToLeft**|
| **xlToRight**|
| **xlUp**|
If the  **[MoveAfterReturn](application-moveafterreturn-property-excel.md)** property is **False** , the selection doesn't move at all, regardless of how the **MoveAfterReturnDirection** property is set.


## Example

This example causes the active cell to move to the right when the user presses ENTER.


```vb
Application.MoveAfterReturn = True 
Application.MoveAfterReturnDirection = xlToRight
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

