---
title: Workbook.SaveAsXMLData Method (Excel)
keywords: vbaxl10.chm199232
f1_keywords:
- vbaxl10.chm199232
ms.prod: excel
api_name:
- Excel.Workbook.SaveAsXMLData
ms.assetid: 7c4c1be3-d3a5-6e90-7750-9f371f008541
ms.date: 06/08/2017
---


# Workbook.SaveAsXMLData Method (Excel)

Exports the data that has been mapped to the specified XML schema map to an XML data file.


## Syntax

 _expression_ . **SaveAsXMLData**( **_Filename_** , **_Map_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Required| **String**|A string that indicates the name of the file to be saved. You can include a full path; if you don't, Microsoft Excel saves the file in the current folder.|
| _Map_|Required| **[XmlMap](xmlmap-object-excel.md)**|The schema map to apply to the data.|

## Remarks

This method will result in a run-time error if Excel cannot export data with the specified schema map. To check whether Excel can use the specified schema map to export data, use the  **[IsExportable](xmlmap-isexportable-property-excel.md)** property.


## Example

The following example verifies that Excel can use the schema map "Customer" to export data, and then exports the data mapped to the "Customer" schema map to a file named "Customer Data.xml".


```vb
Sub ExportAsXMLData() 
 Dim objMapToExport As XmlMap 
 
 Set objMapToExport = ActiveWorkbook.XmlMaps("Customer") 
 
 If objMapToExport.IsExportable Then 
 
 ActiveWorkbook.SaveAsXMLData "Customer Data.xml", objMapToExport 
 Else 
 MsgBox "Cannot use " &; objMapToExport.Name &; _ 
 "to export the contents of the worksheet to XML data." 
 End If 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

