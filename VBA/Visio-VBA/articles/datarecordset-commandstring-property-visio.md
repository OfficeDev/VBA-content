---
title: DataRecordset.CommandString Property (Visio)
keywords: vis_sdr.chm16460300
f1_keywords:
- vis_sdr.chm16460300
ms.prod: visio
api_name:
- Visio.DataRecordset.CommandString
ms.assetid: 7d9151b0-db8c-a8ce-edea-7ef25d241e98
ms.date: 06/08/2017
---


# DataRecordset.CommandString Property (Visio)

Gets or sets the command string used to query the data source to create a data recordset or refresh an existing one. Read/write.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **CommandString**

 _expression_ An expression that returns a **DataRecordset** object.


### Return Value

String


## Remarks

The command string of a data recordset specifies the database table or Microsoft Excel worksheet and the columns within the table or worksheet that contain the data you want to query. The command string is also passed to the  **[DataRecordset.Refresh](datarecordset-refresh-method-visio.md)** method when the data in the data recordset is refreshed.

Setting the  **CommandString** property to a new value has no effect on data already in the data recordset. To update the data in a data recordset using a new **CommandString** setting, call the **Refresh** method.

The  **CommandString** property does not apply to data recordsets created by using the **[AddFromXML](datarecordsets-addfromxml-method-visio.md)** method.

The following sample command string directs Visio to retrieve all data from an Excel worksheet named Sheet1:  `"SELECT * FROM [Sheet1$]"`.


