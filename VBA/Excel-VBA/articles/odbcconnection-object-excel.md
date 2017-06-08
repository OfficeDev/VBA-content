---
title: ODBCConnection Object (Excel)
keywords: vbaxl10.chm795072
f1_keywords:
- vbaxl10.chm795072
ms.prod: excel
api_name:
- Excel.ODBCConnection
ms.assetid: b880ebec-15a4-5a3d-ef02-db73106db9c9
ms.date: 06/08/2017
---


# ODBCConnection Object (Excel)

Represents the ODBC connection.


## Remarks

An ODBC connection can be stored in an Excel workbook. When Microsoft Excel opens the workbook, Excel creates an in-memory copy of the ODBC connection known as the  **ODBCConnection** object.

An  **ODBCConnection** object contains information related to the connection, such as the name of the server to connect to and the name of the objects to be opened on that server. Optionally, the **ODBCConnection** object may also include authentication credential information, or a command that is to be passed to the server and executed (for example, a `SELECT` statement to be executed by SQL Server).


## Methods



|**Name**|
|:-----|
|[CancelRefresh](odbcconnection-cancelrefresh-method-excel.md)|
|[Refresh](odbcconnection-refresh-method-excel.md)|
|[SaveAsODC](odbcconnection-saveasodc-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[AlwaysUseConnectionFile](odbcconnection-alwaysuseconnectionfile-property-excel.md)|
|[Application](odbcconnection-application-property-excel.md)|
|[BackgroundQuery](odbcconnection-backgroundquery-property-excel.md)|
|[CommandText](odbcconnection-commandtext-property-excel.md)|
|[CommandType](odbcconnection-commandtype-property-excel.md)|
|[Connection](odbcconnection-connection-property-excel.md)|
|[Creator](odbcconnection-creator-property-excel.md)|
|[EnableRefresh](odbcconnection-enablerefresh-property-excel.md)|
|[Parent](odbcconnection-parent-property-excel.md)|
|[RefreshDate](odbcconnection-refreshdate-property-excel.md)|
|[Refreshing](odbcconnection-refreshing-property-excel.md)|
|[RefreshOnFileOpen](odbcconnection-refreshonfileopen-property-excel.md)|
|[RefreshPeriod](odbcconnection-refreshperiod-property-excel.md)|
|[RobustConnect](odbcconnection-robustconnect-property-excel.md)|
|[SavePassword](odbcconnection-savepassword-property-excel.md)|
|[ServerCredentialsMethod](odbcconnection-servercredentialsmethod-property-excel.md)|
|[ServerSSOApplicationID](odbcconnection-serverssoapplicationid-property-excel.md)|
|[SourceConnectionFile](odbcconnection-sourceconnectionfile-property-excel.md)|
|[SourceData](odbcconnection-sourcedata-property-excel.md)|
|[SourceDataFile](odbcconnection-sourcedatafile-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
