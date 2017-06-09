---
title: DataConnection.FileName Property (Visio)
keywords: vis_sdr.chm16560380
f1_keywords:
- vis_sdr.chm16560380
ms.prod: visio
api_name:
- Visio.DataConnection.FileName
ms.assetid: fd8fb240-e9b8-05d9-fb59-8e9d412ca346
ms.date: 06/08/2017
---


# DataConnection.FileName Property (Visio)

Gets the name of the Office Data Connection (ODC) file that contains the connection string and query command string for the data connection. Read-only.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **FileName**

 _expression_ An expression that returns a **DataConnection** object.


### Return Value

String


## Remarks

An ODC file contains a connection string that specifies how to connect to an OLEDB or ODBC data source and a query command string that specifies how to extract the desired data from the data source. An ODC file uses HTML and XML to store connection and query information. You can view or edit the contents of the file in any text editor. ODC files have the .odc file name extension. 

When you use the  **[DataRecordsets.AddFromConnectionFile](datarecordsets-addfromconnectionfile-method-visio.md)** to create a new data recordset, you pass the method a string pointing to an ODC file as the FileName parameter. That string then becomes the value of the **FileName** property for the **DataConnection** associated with the resulting data recordset. If the **DataConnection** object is not associated with an ODC file, the **FileName** property returns the name and full path of the data-source file passed to the **[DataRecordsets.Add](datarecordsets-add-method-visio.md)** method as the value of the "Data Source" attribute of the ConnectionString parameter when the data recordset was created.

You can use the  **Data Connection Wizard** in Microsoft Access or Microsoft Excel to create an ODC file that will connect to and retrieve the data you want.


