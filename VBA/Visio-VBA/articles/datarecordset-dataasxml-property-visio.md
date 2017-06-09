---
title: DataRecordset.DataAsXML Property (Visio)
keywords: vis_sdr.chm16460305
f1_keywords:
- vis_sdr.chm16460305
ms.prod: visio
api_name:
- Visio.DataRecordset.DataAsXML
ms.assetid: 500dda1a-0747-57d0-f847-e3e1f72e96a3
ms.date: 06/08/2017
---


# DataRecordset.DataAsXML Property (Visio)

Returns an XML string that fully describes a data recordset and conforms to the Microsoft ActiveX® Data Objects (ADO) classic XML schema. Read-only.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **DataAsXML**

 _expression_ An expression that returns a **DataRecordset** object.


### Return Value

String


## Remarks

The string returned by  **DataAsXML** contains all the rows in the data recordset with Microsoft Visio row IDs pre-pended to them.

The string returned by  **DataAsXML** contains all the valid rows and columns in the data recordset that was imported as well as an additional column, named _Visio_RowID_, inserted as the first column, that assigns a unique row ID to each row in the data recordset.


