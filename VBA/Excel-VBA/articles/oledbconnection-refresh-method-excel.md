---
title: OLEDBConnection.Refresh Method (Excel)
keywords: vbaxl10.chm794083
f1_keywords:
- vbaxl10.chm794083
ms.prod: excel
api_name:
- Excel.OLEDBConnection.Refresh
ms.assetid: c28e9443-81e2-dfec-a3fb-a127c3fa2918
ms.date: 06/08/2017
---


# OLEDBConnection.Refresh Method (Excel)

Refreshes an OLE DB connection.


## Syntax

 _expression_ . **Refresh**

 _expression_ A variable that represents an **OLEDBConnection** object.


## Remarks

When making the connection to the OLE DB data source, Microsoft Excel uses the connection string specified by the  **[Connection](oledbconnection-connection-property-excel.md)** property. If the specified connection string is missing required values, dialog boxes will be displayed to prompt the user for the required information. If the **[DisplayAlerts](application-displayalerts-property-excel.md)** property is **False** , dialog boxes are not displayed and the **Refresh** method fails with the Insufficient Connection Information exception.

After Microsoft Excel makes a successful connection, it stores the completed connection string so that prompts will not be displayed for subsequent calls to the  **Refresh** method during the same editing session. You can obtain the completed connection string by examining the value of the **[Connection](oledbconnection-connection-property-excel.md)** property.

After the database connection is made, the SQL query is validated. If the query is not valid, the  **Refresh** method fails with the SQL Syntax Error exception.

The  **Refresh** method returns **True** if the query is successfully completed or started; it returns **False** if the user cancels a connection.

To see whether the number of fetched rows exceeded the number of available rows on the worksheet, examine the  **[FetchedRowOverflow](querytable-fetchedrowoverflow-property-excel.md)** property. This property is initialized every time the **Refresh** method is called.


## See also


#### Concepts


[OLEDBConnection Object](oledbconnection-object-excel.md)

