---
title: WorkbookConnection Object (Excel)
keywords: vbaxl10.chm773072
f1_keywords:
- vbaxl10.chm773072
ms.prod: excel
api_name:
- Excel.WorkbookConnection
ms.assetid: 5974dd57-7671-cd55-3f8f-6a76fa938317
ms.date: 06/08/2017
---


# WorkbookConnection Object (Excel)

A connection is a set of information needed to obtain data from an external data source other than an Microsoft Office Excel 2007 workbook. 


## Remarks

Connections can be stored within an Excel workbook. When the workbook is opened, Excel creates an in-memory copy of the connection which is refered to as the connection object. A connection object contains information such as the name of the server and the name of the object to be opened on that server. Optionally, the connection object may also include authentication credentials and/or a command that is to be passed to the server and executed (example: a SELECT statement to be executed by SQL Server.) 


 **Note**  The exact form of the connection depends on the mechanism that is being used to retrieve data - ODBC connections, OLEDB connections, and Web Queries will contain different information.

A connection may also be stored in a separate connection file. Most connections in an Excel workbook include a pointer to an external connection file. Connection files have extensions that clearly label them as connection files (*.ODC, *.IQY, etc.) and may be located on the user's local machine or in other well known or trusted locations such as WSS (Data Connection Library), or other corporate servers. Connection files enable multiple users within the same organization to re-use connections. Network administrators are able to change the way the entire organization connects to a back-end data source by changing a single connection file. A connection file is not always required when connecting to an external data source.

Connection names are strings that uniquely identify connections within the workbook in which they are used. There are other properties of a connection that are not unique. Whenever a formula in Excel takes an argument that is a connection, it will be sufficient to refer to the name of that connection, either directly (as a string) or indirectly (by referring to a cell that contains the connection name as a string.)


## Methods



|**Name**|
|:-----|
|[Delete](workbookconnection-delete-method-excel.md)|
|[Refresh](workbookconnection-refresh-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](workbookconnection-application-property-excel.md)|
|[Creator](workbookconnection-creator-property-excel.md)|
|[DataFeedConnection](workbookconnection-datafeedconnection-property-excel.md)|
|[Description](workbookconnection-description-property-excel.md)|
|[InModel](workbookconnection-inmodel-property-excel.md)|
|[ModelConnection](workbookconnection-modelconnection-property-excel.md)|
|[ModelTables](workbookconnection-modeltables-property-excel.md)|
|[Name](workbookconnection-name-property-excel.md)|
|[ODBCConnection](workbookconnection-odbcconnection-property-excel.md)|
|[OLEDBConnection](workbookconnection-oledbconnection-property-excel.md)|
|[Parent](workbookconnection-parent-property-excel.md)|
|[Ranges](workbookconnection-ranges-property-excel.md)|
|[RefreshWithRefreshAll](workbookconnection-refreshwithrefreshall-property-excel.md)|
|[TextConnection](workbookconnection-textconnection-property-excel.md)|
|[Type](workbookconnection-type-property-excel.md)|
|[WorksheetDataConnection](workbookconnection-worksheetdataconnection-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
