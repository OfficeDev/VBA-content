---
title: DataColumn.Name Property (Visio)
keywords: vis_sdr.chm16713930
f1_keywords:
- vis_sdr.chm16713930
ms.prod: visio
api_name:
- Visio.DataColumn.Name
ms.assetid: 9b66e7b1-f3e7-af3a-8a64-59ab90fd6119
ms.date: 06/08/2017
---


# DataColumn.Name Property (Visio)

Gets the unique name of the data column in its parent data recordset. Read-only.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **Name**

 _expression_ A variable that represents a **DataColumn** object.


### Return Value

String


## Remarks

The  **Name** property is the default property of the **DataColumn** object. The value of **Name** is unique within a particular data recordset and therefore uniquely identifies the column in the data recordset. The value that Visio assigns for the **Name** property is the same as, or derived from, the name of the column in the original data source.

For a given column, the value of the  **Name** property is not necessarily the same as that of the **[DisplayName](datacolumn-displayname-property-visio.md)** property, which specifies the name of the column in the **External Data** window in the Visio user interface, and which is read/write.


