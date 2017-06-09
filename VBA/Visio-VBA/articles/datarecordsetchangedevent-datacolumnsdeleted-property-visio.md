---
title: DataRecordsetChangedEvent.DataColumnsDeleted Property (Visio)
keywords: vis_sdr.chm17260470
f1_keywords:
- vis_sdr.chm17260470
ms.prod: visio
api_name:
- Visio.DataRecordsetChangedEvent.DataColumnsDeleted
ms.assetid: 6fae59a1-cacc-076f-fd9d-1efbf5f1972e
ms.date: 06/08/2017
---


# DataRecordsetChangedEvent.DataColumnsDeleted Property (Visio)

After data in a data recordset are refreshed, returns an array of names of data columns deleted from the data recordset as a result of the refresh operation. Read-only.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **DataColumnsDeleted**

 _expression_ An expression that returns a **DataRecordsetChangedEvent** object.


### Return Value

String()


## Remarks

The columns returned by this property have already been deleted. As a result, you can no longer use Visio Automation properties or methods to retrieve the  **DataColumn** objects that represented the columns, nor any information about data formerly in these columns.


