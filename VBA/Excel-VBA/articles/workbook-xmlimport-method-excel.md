---
title: Workbook.XmlImport Method (Excel)
keywords: vbaxl10.chm199226
f1_keywords:
- vbaxl10.chm199226
ms.prod: excel
api_name:
- Excel.Workbook.XmlImport
ms.assetid: 97964c82-1fbe-7060-0a90-23c190e0b398
ms.date: 06/08/2017
---


# Workbook.XmlImport Method (Excel)

Imports an XML data file into the current workbook.


## Syntax

 _expression_ . **XmlImport**( **_Url_** , **_ImportMap_** , **_Overwrite_** , **_Destination_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Url_|Required| **String**|A uniform resource locator (URL) or a uniform naming convention (UNC) path to a XML data file.|
| _ImportMap_|Required| **[XmlMap](xmlmap-object-excel.md)**|The schema map to apply when importing the file. If the data was previously imported, contains a reference to the  **XmlMap** object containing the data.|
| _Overwrite_|Optional| **Variant**|If a value is not specified for the Destination parameter, then this parameter specifies whether or not to overwrite data that has been mapped to the schema map specified in the ImportMap parameter. Set to  **True** to overwrite the data or **False** to append the new data to the existing data. The default value is **True** . If a value is specified for the Destination parameter, then this parameter specifies whether or not to overwrite existing data. Set to **True** to overwrite existing data or **False** to cancel the import if data would be overwritten. The default value is **True** .|
| _Destination_|Optional| **Variant**|Specifies the range where the list will be created. You only use the top left corner of the range.|

### Return Value

[XlXmlImportResult](xlxmlimportresult-enumeration-excel.md)


## Remarks



| **XlXmlImportResult** can be one of the following **XlXmlImportResult** constants.|
| **xlXmlImportElementsTruncated** . The contents of the specified XML data file have been truncated because the XML data file is too large for the worksheet.|
| **xlXmlImportSuccess** . The XML data file was successfully imported.|
| **xlXmlImportValidationFailed** . The data being imported failed schema validation, but was imported anyway.|
This method allows you to import data into the workbook from a file path. Excel uses the first qualifying map found, or if the destination range is specified, Excel will automatically list the data.

Don't specify a value for the  _Destination_ parameter if you want to import data into an existing mapping.

The following conditions will cause the  **XMLImport** method to generate run-time errors:


- The specified XML data contains syntax errors.
    
- The import process was cancelled because the specified data cannot fit into the worksheet.
    
- If no qualifying maps are found and the destination range was not specified.
    


Use the  **[XmlImportXml](workbook-xmlimportxml-method-excel.md)** method of the **[Workbook](workbook-object-excel.md)** object to import XML data that has been previously loaded into memory.


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

