---
title: QueryTable.RobustConnect Property (Excel)
keywords: vbaxl10.chm518133
f1_keywords:
- vbaxl10.chm518133
ms.prod: excel
api_name:
- Excel.QueryTable.RobustConnect
ms.assetid: ad180446-82d7-7b5b-59a2-b0de299ae934
ms.date: 06/08/2017
---


# QueryTable.RobustConnect Property (Excel)

Returns or sets how the query table connects to its data source. Read/write  **[XlRobustConnect](xlrobustconnect-enumeration-excel.md)** .


## Syntax

 _expression_ . **RobustConnect**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks



| **XlRobustConnect** can be one of these **XlRobustConnect** constants.|
| **xlAlways** . The query table always uses external source information (as defined by the **[SourceConnectionFile](querytable-sourceconnectionfile-property-excel.md)** or **[SourceDataFile](querytable-sourcedatafile-property-excel.md)** property) to reconnect.|
| **xlAsRequired** . The query table uses external source information to reconnect, using the **[Connection](querytable-connection-property-excel.md)** property.|
| **xlNever** . The query table never uses source information to reconnect.|
If you import data by using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data by using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **RobustConnect** property.


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

