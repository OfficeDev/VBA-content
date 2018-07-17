---
title: QueryTable.SaveData Property (Excel)
keywords: vbaxl10.chm518095
f1_keywords:
- vbaxl10.chm518095
ms.prod: excel
api_name:
- Excel.QueryTable.SaveData
ms.assetid: 7657e1ee-cbed-91c6-0e69-defe4ca69897
ms.date: 06/08/2017
---


# QueryTable.SaveData Property (Excel)

 **True** if data for the QueryTable report is saved with the workbook. **False** if only the report definition is saved. Read/write **Boolean** .


## Syntax

 _expression_ . **SaveData**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

For OLAP data sources, this property is always set to  **False** .

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **SaveData** property.


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

