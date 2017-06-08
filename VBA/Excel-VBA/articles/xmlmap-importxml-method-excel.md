---
title: XmlMap.ImportXml Method (Excel)
keywords: vbaxl10.chm754088
f1_keywords:
- vbaxl10.chm754088
ms.prod: excel
api_name:
- Excel.XmlMap.ImportXml
ms.assetid: 07db07d3-cd0f-08fe-3463-04ca72d084d1
ms.date: 06/08/2017
---


# XmlMap.ImportXml Method (Excel)

Imports XML data from a  **String** variable into cells that have been mapped to the specified **[XmlMap](xmlmap-object-excel.md)** object.


## Syntax

 _expression_ . **ImportXml**( **_XmlData_** , **_Overwrite_** )

 _expression_ A variable that represents a **XmlMap** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XmlData_|Required| **String**|The string that contains the XML data to import.|
| _Overwrite_|Optional| **Variant**|Specifies whether to overwrite the contents of cells that are currently mapped to the specified XML map. Set to  **True** to overwrite the cells; set to **False** to append the data to the existing range. If this parameter is not specified, the current value of the **[AppendOnImport](xmlmap-appendonimport-property-excel.md)** property of the XML map determines whether the contents of cells are overwritten or not.|

### Return Value

[XlXmlImportResult](xlxmlimportresult-enumeration-excel.md)


## Remarks



| **XlXmlImportResult** can be one of the following **XlXmlImportResult** constants.|
| **xlXmlImportElementsTruncated** . The contents of the specified XML data file have been truncated because the XML data file is too large for the worksheet.|
| **xlXmlImportSuccess** . The XML data file was successfully imported.|
| **xlXmlImportValidationFailed** . The data being imported failed schema validation, but was imported anyway.|
To import the contents of an XML data file into cells mapped to a specific schema map, use the  **[Import](xmlmap-import-method-excel.md)** method of the **XmlMap** object.

If either of the following conditions is true, a runtime error will occur. If more than one condition is true, Excel returns a runtime error for the most severe (they are listed below with the most severe listed first):


- If the XML data contains syntactical errors.
    
- If import is cancelled because not all of the data could fit in the worksheet.
    

## See also


#### Concepts


[XmlMap Object](xmlmap-object-excel.md)

