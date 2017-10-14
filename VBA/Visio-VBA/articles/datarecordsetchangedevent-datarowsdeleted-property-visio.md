---
title: DataRecordsetChangedEvent.DataRowsDeleted Property (Visio)
keywords: vis_sdr.chm17260460
f1_keywords:
- vis_sdr.chm17260460
ms.prod: visio
api_name:
- Visio.DataRecordsetChangedEvent.DataRowsDeleted
ms.assetid: 9b2b0b6e-702a-824b-ff83-210de5c8c889
ms.date: 06/08/2017
---


# DataRecordsetChangedEvent.DataRowsDeleted Property (Visio)

Returns an array of IDs of data rows deleted from the data recordset as a result of a data-refresh operation. Read-only.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **DataRowsDeleted**

 _expression_ An expression that returns a **DataRecordsetChangedEvent** object.


### Return Value

Long()


## Remarks

The rows returned by this property have already been deleted. As a result, you can no longer retrieve information about data in these rows by calling Visio Automation properties or methods.


