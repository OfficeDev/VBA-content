---
title: DataRecordset.TimeRefreshed Property (Visio)
keywords: vis_sdr.chm16460335
f1_keywords:
- vis_sdr.chm16460335
ms.prod: visio
api_name:
- Visio.DataRecordset.TimeRefreshed
ms.assetid: ebdf1acd-81f9-bd5e-48ba-d34100a8f702
ms.date: 06/08/2017
---


# DataRecordset.TimeRefreshed Property (Visio)

Returns the date and time the data recordset was last refreshed. Read-only.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **TimeRefreshed**

 _expression_ An expression that returns a **DataRecordset** object.


### Return Value

Date


## Remarks

The  **TimeRefreshed** property value is returned in Coordinated Universal Time (Greenwich Mean Time).

If you successfully create a data recordset but it fails to retrieve any data from the data source,  **TimeRefreshed** returns zero.

The first time you execute a query against a data recordset,  **TimeRefreshed** is set to the time the query is executed. If, subsequently, the **[Refresh](datarecordset-refresh-method-visio.md)** method is called, **TimeRefreshed** is set to the time the data recordset is refreshed.


