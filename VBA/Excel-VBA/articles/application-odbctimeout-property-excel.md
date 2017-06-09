---
title: Application.ODBCTimeout Property (Excel)
keywords: vbaxl10.chm133175
f1_keywords:
- vbaxl10.chm133175
ms.prod: excel
api_name:
- Excel.Application.ODBCTimeout
ms.assetid: 92262209-6a0f-f58f-e2d7-2f502f6bd397
ms.date: 06/08/2017
---


# Application.ODBCTimeout Property (Excel)

Returns or sets the ODBC query time limit, in seconds. The default value is 45 seconds. Read/write  **Long** .


## Syntax

 _expression_ . **ODBCTimeout**

 _expression_ A variable that represents an **Application** object.


## Remarks

The value 0 (zero) indicates an indefinite time limit.


## Example

This example sets the ODBC query time limit to 15 seconds.


```vb
Application.ODBCTimeout = 15
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

