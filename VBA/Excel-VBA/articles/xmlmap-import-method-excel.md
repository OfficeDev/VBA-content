---
title: XmlMap.Import Method (Excel)
keywords: vbaxl10.chm754087
f1_keywords:
- vbaxl10.chm754087
ms.prod: excel
api_name:
- Excel.XmlMap.Import
ms.assetid: 60265bbd-4994-8fba-7072-ec5dada885d3
ms.date: 06/08/2017
---


# XmlMap.Import Method (Excel)

Imports data from the specified XML data file into cells that have been mapped to the specified  **[XmlMap](xmlmap-object-excel.md)** object.


## Syntax

 _expression_ . **Import**( **_Url_** , **_Overwrite_** )

 _expression_ A variable that represents a **XmlMap** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Url_|Required| **String**|The path to the XML data to import. The path can be specified in Universal Naming convention (UNC) or Uniform Resource Locator (URL) format. The file can be an XML data file.|
| _Overwrite_|Optional| **Variant**|Set to  **True** to overwrite existing data. Set to **False** to append to existing data. The default value is **False** .|

### Return Value

A  **[XlXmlImportResult](xlxmlimportresult-enumeration-excel.md)** value that indicates the result of the method.


## Remarks

This method returns one of the following  **XlXmlImportResult** constants:



| **xlXmlImportElementsTruncated** The contents of the specified XML data file have been truncated because the XML data file is too large for the worksheet.|
| **xlXmlImportSuccess** The XML data file was successfully imported.|
| **xlXmlImportValidationFailed** The data being imported failed schema validation, but was imported anyway.|
If either of the following conditions is true, a runtime error will occur. If more than one condition is true, Excel returns a runtime error for the most severe (they are listed below with the most severe listed first):


- If the XML data contains syntactical errors.
    
- If import is cancelled because not all of the data could fit in the worksheet.
    

## See also


#### Concepts


[XmlMap Object](xmlmap-object-excel.md)

