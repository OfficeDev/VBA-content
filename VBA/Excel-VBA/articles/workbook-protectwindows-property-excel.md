---
title: Workbook.ProtectWindows Property (Excel)
keywords: vbaxl10.chm199132
f1_keywords:
- vbaxl10.chm199132
ms.prod: excel
api_name:
- Excel.Workbook.ProtectWindows
ms.assetid: 0f285fbe-2545-5c7d-9e3d-f08d57e78092
ms.date: 06/08/2017
---


# Workbook.ProtectWindows Property (Excel)

 **True** if the windows of the workbook are protected. Read-only **Boolean** .


## Syntax

 _expression_ . **ProtectWindows**

 _expression_ A variable that represents a **Workbook** object.


## Example

This example displays a message if the windows in the active workbook are protected.


```vb
If ActiveWorkbook.ProtectWindows = True Then 
 MsgBox "Remember, you cannot rearrange any" &; _ 
 " window in this workbook." 
End If 

```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

