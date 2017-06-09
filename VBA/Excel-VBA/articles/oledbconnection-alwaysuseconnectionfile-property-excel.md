---
title: OLEDBConnection.AlwaysUseConnectionFile Property (Excel)
keywords: vbaxl10.chm794099
f1_keywords:
- vbaxl10.chm794099
ms.prod: excel
api_name:
- Excel.OLEDBConnection.AlwaysUseConnectionFile
ms.assetid: de9cd9a7-0dd6-7ee2-d48f-bd61a7006c1e
ms.date: 06/08/2017
---


# OLEDBConnection.AlwaysUseConnectionFile Property (Excel)

 **True** if the connection file is always used to establish connection to the data source. Read/write **Boolean** .


## Syntax

 _expression_ . **AlwaysUseConnectionFile**

 _expression_ A variable that represents an **OLEDBConnection** object.


## Remarks

When this property is  **True** the connection file will always be used to establish the connection to the data source. If the connection embedded within the workbook is different from the external connection file, the embedded connection will be ignored and the external connection file will be the only version considered.


## See also


#### Concepts


[OLEDBConnection Object](oledbconnection-object-excel.md)

