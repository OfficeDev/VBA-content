---
title: Workbook.XmlImportXml Method (Excel)
keywords: vbaxl10.chm199231
f1_keywords:
- vbaxl10.chm199231
ms.prod: excel
api_name:
- Excel.Workbook.XmlImportXml
ms.assetid: b0edbe49-f578-ead0-8371-0196f5c515d4
ms.date: 06/08/2017
---


# Workbook.XmlImportXml Method (Excel)

Imports an XML data stream that has been previously loaded into memory. Excel uses the first qualifying map found or if the destination range is specified, Excel will automatically list the data.


## Syntax

 _expression_ . **XmlImportXml**( **_Data_** , **_ImportMap_** , **_Overwrite_** , **_Destination_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Data_|Required| **String**|The data to import.|
| _ImportMap_|Required| **[XmlMap](xmlmap-object-excel.md)**|The schema map to apply when importing the file.|
| _Overwrite_|Optional| **Variant**|If a value is not specified for the Destination parameter, then this parameter specifies whether or not to overwrite data that has been mapped to the schema map specified in the ImportMap parameter. Set to  **True** to overwrite the data or **False** to append the new data to the existing data. The default value is **True** . If a value is specified for the Destination parameter, then this parameter specifies whether or not to overwrite existing data. Set to **True** to overwrite existing data or **False** to cancel the import if data would be overwritten. The default value is **True** .|
| _Destination_|Optional| **Variant**|Specifies the range where the list will be created. Excel only uses the top left corner of the range.|

### Return Value

[XlXmlImportResult](xlxmlimportresult-enumeration-excel.md)


## Remarks



| **XlXmlImportResult** can be one of the following **XlXmlImportResult** constants.|
| **xlXmlImportElementsTruncated** . The contents of the specified XML data file have been truncated because the XML data file is too large for the worksheet.|
| **xlXmlImportSuccess** . The XML data file was successfully imported.|
| **xlXmlImportValidationFailed** . The data being imported failed schema validation, but was imported anyway.|
Don't specify a value for the  _Destination_ parameter if you want to import data into an existing mapping.

The following conditions will cause the  **[XmlImport](workbook-xmlimport-method-excel.md)** method to generate run-time errors:


- The specified XML data contains syntax errors.
    
- The import process was cancelled because the specified data cannot fit into the worksheet.
    
- If no qualifying maps are found and the destination range was not specified.
    


Use the  **XMLImport** method of the **[Workbook](workbook-object-excel.md)** object to import an XML data file into the current workbook.


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

