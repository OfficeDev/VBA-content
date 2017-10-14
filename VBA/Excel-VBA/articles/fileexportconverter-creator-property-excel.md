---
title: FileExportConverter.Creator Property (Excel)
keywords: vbaxl10.chm862074
f1_keywords:
- vbaxl10.chm862074
ms.prod: excel
api_name:
- Excel.FileExportConverter.Creator
ms.assetid: f008a8c9-89a6-a0a9-4f26-acffdde29e6a
ms.date: 06/08/2017
---


# FileExportConverter.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[FileExportConverter](fileexportconverter-object-excel.md)** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the hexadecimal value, 5843454C, which represents the string "XCEL". The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


[FileExportConverter Object](fileexportconverter-object-excel.md)

