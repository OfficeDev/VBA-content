---
title: OLEDBConnection.RetrieveInOfficeUILang Property (Excel)
keywords: vbaxl10.chm794104
f1_keywords:
- vbaxl10.chm794104
ms.prod: excel
api_name:
- Excel.OLEDBConnection.RetrieveInOfficeUILang
ms.assetid: 51d2a8b7-75e6-c503-895b-0f5ab8d66265
ms.date: 06/08/2017
---


# OLEDBConnection.RetrieveInOfficeUILang Property (Excel)

 **True** if the data and errors are to be retrieved in the Office user interface display language when available. Read/write **Boolean** .


## Syntax

 _expression_ . **RetrieveInOfficeUILang**

 _expression_ A variable that represents an **OLEDBConnection** object.


## Remarks

If this property is set to  **False** , the LCID value in the connection string is used instead. If an LCID is not specified, the default LCID from the server is used.


## See also


#### Concepts


[OLEDBConnection Object](oledbconnection-object-excel.md)

