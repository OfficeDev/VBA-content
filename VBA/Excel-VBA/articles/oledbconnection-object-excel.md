---
title: OLEDBConnection Object (Excel)
keywords: vbaxl10.chm793072
f1_keywords:
- vbaxl10.chm793072
ms.prod: excel
api_name:
- Excel.OLEDBConnection
ms.assetid: f246e544-9854-8e71-a7f7-dec57dd725e4
ms.date: 06/08/2017
---


# OLEDBConnection Object (Excel)

Represents the OLE DB connection.


## Remarks

An OLE DB connection can be stored in an Excel workbook. When Micrososft Excel opens the workbook, Excel creates an in-memory copy of the OLE DB connection known as the  **OLEDBConnection** object.

An  **OLEDBConnection** object contains information related to the connection, such as the name of the server to connect to and the name of the objects to be opened on that server. Optionally, the **OLEDBConnection** object may also include authentication credential information, or a command that is to be passed to the server and executed (for example, a `SELECT` statement to be executed by SQL Server).


## Methods



|**Name**|
|:-----|
|[CancelRefresh](oledbconnection-cancelrefresh-method-excel.md)|
|[MakeConnection](oledbconnection-makeconnection-method-excel.md)|
|[Reconnect](oledbconnection-reconnect-method-excel.md)|
|[Refresh](oledbconnection-refresh-method-excel.md)|
|[SaveAsODC](oledbconnection-saveasodc-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[ADOConnection](oledbconnection-adoconnection-property-excel.md)|
|[AlwaysUseConnectionFile](oledbconnection-alwaysuseconnectionfile-property-excel.md)|
|[Application](oledbconnection-application-property-excel.md)|
|[BackgroundQuery](oledbconnection-backgroundquery-property-excel.md)|
|[CalculatedMembers](oledbconnection-calculatedmembers-property-excel.md)|
|[CommandText](oledbconnection-commandtext-property-excel.md)|
|[CommandType](oledbconnection-commandtype-property-excel.md)|
|[Connection](oledbconnection-connection-property-excel.md)|
|[Creator](oledbconnection-creator-property-excel.md)|
|[EnableRefresh](oledbconnection-enablerefresh-property-excel.md)|
|[IsConnected](oledbconnection-isconnected-property-excel.md)|
|[LocalConnection](oledbconnection-localconnection-property-excel.md)|
|[LocaleID](oledbconnection-localeid-property-excel.md)|
|[MaintainConnection](oledbconnection-maintainconnection-property-excel.md)|
|[MaxDrillthroughRecords](oledbconnection-maxdrillthroughrecords-property-excel.md)|
|[OLAP](oledbconnection-olap-property-excel.md)|
|[Parent](oledbconnection-parent-property-excel.md)|
|[RefreshDate](oledbconnection-refreshdate-property-excel.md)|
|[Refreshing](oledbconnection-refreshing-property-excel.md)|
|[RefreshOnFileOpen](oledbconnection-refreshonfileopen-property-excel.md)|
|[RefreshPeriod](oledbconnection-refreshperiod-property-excel.md)|
|[RetrieveInOfficeUILang](oledbconnection-retrieveinofficeuilang-property-excel.md)|
|[RobustConnect](oledbconnection-robustconnect-property-excel.md)|
|[SavePassword](oledbconnection-savepassword-property-excel.md)|
|[ServerCredentialsMethod](oledbconnection-servercredentialsmethod-property-excel.md)|
|[ServerFillColor](oledbconnection-serverfillcolor-property-excel.md)|
|[ServerFontStyle](oledbconnection-serverfontstyle-property-excel.md)|
|[ServerNumberFormat](oledbconnection-servernumberformat-property-excel.md)|
|[ServerSSOApplicationID](oledbconnection-serverssoapplicationid-property-excel.md)|
|[ServerTextColor](oledbconnection-servertextcolor-property-excel.md)|
|[SourceConnectionFile](oledbconnection-sourceconnectionfile-property-excel.md)|
|[SourceDataFile](oledbconnection-sourcedatafile-property-excel.md)|
|[UseLocalConnection](oledbconnection-uselocalconnection-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
