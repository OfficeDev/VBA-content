---
title: ODBCConnection.Refresh Method (Excel)
keywords: vbaxl10.chm796079
f1_keywords:
- vbaxl10.chm796079
ms.prod: excel
api_name:
- Excel.ODBCConnection.Refresh
ms.assetid: 26a9ba46-1679-f83b-6933-b6c448dce9e7
ms.date: 06/08/2017
---


# ODBCConnection.Refresh Method (Excel)

Refreshes an ODBC connection.


## Syntax

 _expression_ . **Refresh**

 _expression_ A variable that represents an **ODBCConnection** object.


## Remarks

When making the connection to the ODBC data source, Microsoft Excel uses the connection string specified by the  **[Connection](odbcconnection-connection-property-excel.md)** property. If the specified connection string is missing required values, dialog boxes will be displayed to prompt the user for the required information. If the **[DisplayAlerts](application-displayalerts-property-excel.md)** property is **False** , dialog boxes are not displayed and the **Refresh** method fails with the Insufficient Connection Information exception.

After Microsoft Excel makes a successful connection, it stores the completed connection string so that prompts will not be displayed for subsequent calls to the  **Refresh** method during the same editing session. You can obtain the completed connection string by examining the value of the **[Connection](odbcconnection-connection-property-excel.md)** property.

After the database connection is made, the SQL query is validated. If the query is not valid, the  **Refresh** method fails with the SQL Syntax Error exception.

If the query requires parameters, the  **[Parameters](parameters-object-excel.md)** collection must be initialized with parameter binding information before the **Refresh** method is called. If not enough parameters have been bound, the **Refresh** method fails with the Parameter Error exception. If parameters are set to prompt for their values, dialog boxes are displayed to the user regardless of the setting of the **[DisplayAlerts](application-displayalerts-property-excel.md)** property. If the user cancels a parameter dialog box, the **Refresh** method halts and returns **False** . If extra parameters are bound with the **Parameters** collection, these extra parameters are ignored.

The  **Refresh** method returns **True** if the query is successfully completed or started; it returns **False** if the user cancels a connection or parameter dialog box.

To see whether the number of fetched rows exceeded the number of available rows on the worksheet, examine the  **[FetchedRowOverflow](querytable-fetchedrowoverflow-property-excel.md)** property. This property is initialized every time the **Refresh** method is called.


## See also


#### Concepts


[ODBCConnection Object](odbcconnection-object-excel.md)

