---
title: FileExportConverters.Creator Property (Excel)
keywords: vbaxl10.chm864074
f1_keywords:
- vbaxl10.chm864074
ms.prod: excel
api_name:
- Excel.FileExportConverters.Creator
ms.assetid: 7310b103-9216-a684-f442-7fd81944b3f5
ms.date: 06/08/2017
---


# FileExportConverters.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[FileExportConverters](fileexportconverters-object-excel.md)** collection.


## Remarks

If the object was created in Microsoft Excel, this property returns the hexadecimal value, 5843454C, which represents the string "XCEL". The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[FileExportConverters Collection](fileexportconverters-object-excel.md)

