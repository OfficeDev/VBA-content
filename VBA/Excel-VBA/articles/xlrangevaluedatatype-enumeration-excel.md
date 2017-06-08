---
title: XlRangeValueDataType Enumeration (Excel)
ms.prod: excel
api_name:
- Excel.XlRangeValueDataType
ms.assetid: a7d50f6e-a1e2-adaf-2516-410210e5fa50
ms.date: 06/08/2017
---


# XlRangeValueDataType Enumeration (Excel)

Specifies the range value data type.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **xlRangeValueDefault**|10|Default. If the specified  **Range** object is empty, returns the value Empty (use the IsEmpty function to test for this case). If the **Range** object contains more than one cell, returns an array of values (use the IsArray function to test for this case).|
| **xlRangeValueMSPersistXML**|12|Returns the recordset representation of the specified  **Range** object in an XML format.|
| **xlRangeValueXMLSpreadsheet**|11|Returns the values, formatting, formulas, and names of the specified  **Range** object in the XML Spreadsheet format.|

