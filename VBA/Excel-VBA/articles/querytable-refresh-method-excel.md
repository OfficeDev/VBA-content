---
title: QueryTable.Refresh Method (Excel)
keywords: vbaxl10.chm518092
f1_keywords:
- vbaxl10.chm518092
ms.prod: excel
api_name:
- Excel.QueryTable.Refresh
ms.assetid: 445d74fb-1a9c-bba4-2d53-0ab0caa876da
ms.date: 06/08/2017
---


# QueryTable.Refresh Method (Excel)

Updates an external data range ( **[QueryTable](querytable-object-excel.md)** ).


## Syntax

 _expression_ . **Refresh**( **_BackgroundQuery_** )

 _expression_ A variable that represents a **QueryTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _BackgroundQuery_|Optional| **Variant**|Used only with  **QueryTables** that are based on the results of a SQL query. **True** to return control to the procedure as soon as a database connection is made and the query is submitted. The **QueryTable** is updated in the background. **False** to return control to the procedure only after all data has been fetched to the worksheet. If this argument isn't specified, the setting of the **[BackgroundQuery](querytable-backgroundquery-property-excel.md)** property determines the query mode.|

### Return Value

Boolean


## Remarks

The following remarks apply to  **QueryTable** objects that are based on the results of a SQL query.

The  **Refresh** method causes Microsoft Excel to connect to the data source of the **QueryTable** object, execute the SQL query, and return data to the range that is based on the **QueryTable** object. Unless this method is called, the **QueryTable** object doesn't communicate with the data source.

When making the connection to the OLE DB or ODBC data source, Microsoft Excel uses the connection string specified by the  **[Connection](querytable-connection-property-excel.md)** property. If the specified connection string is missing required values, dialog boxes will be displayed to prompt the user for the required information. If the **[DisplayAlerts](application-displayalerts-property-excel.md)** property is **False** , dialog boxes aren't displayed and the **Refresh** method fails with the Insufficient Connection Information exception.

After Microsoft Excel makes a successful connection, it stores the completed connection string so that prompts won't be displayed for subsequent calls to the  **Refresh** method during the same editing session. You can obtain the completed connection string by examining the value of the **[Connection](querytable-connection-property-excel.md)** property.

After the database connection is made, the SQL query is validated. If the query isn't valid, the  **Refresh** method fails with the SQL Syntax Error exception.

If the query requires parameters, the  **[Parameters](parameters-object-excel.md)** collection must be initialized with parameter binding information before the **Refresh** method is called. If not enough parameters have been bound, the **Refresh** method fails with the Parameter Error exception. If parameters are set to prompt for their values, dialog boxes are displayed to the user regardless of the setting of the **[DisplayAlerts](application-displayalerts-property-excel.md)** property. If the user cancels a parameter dialog box, the **Refresh** method halts and returns **False** . If extra parameters are bound with the **Parameters** collection, these extra parameters are ignored.

The  **Refresh** method returns **True** if the query is successfully completed or started; it returns **False** if the user cancels a connection or parameter dialog box.

To see whether the number of fetched rows exceeded the number of available rows on the worksheet, examine the  **[FetchedRowOverflow](querytable-fetchedrowoverflow-property-excel.md)** property. This property is initialized every time the **Refresh** method is called.


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

