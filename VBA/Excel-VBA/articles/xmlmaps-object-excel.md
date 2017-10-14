---
title: XmlMaps Object (Excel)
keywords: vbaxl10.chm755072
f1_keywords:
- vbaxl10.chm755072
ms.prod: excel
api_name:
- Excel.XmlMaps
ms.assetid: 0cb16ec8-1120-0da3-508b-c1c9b0aa1701
ms.date: 06/08/2017
---


# XmlMaps Object (Excel)

Represents the collection of  **[XmlMap](xmlmap-object-excel.md)** objects that have been added to a workbook.


## Example

Use the  **[Add](xmlmaps-add-method-excel.md)** method to add an XML map to a workbook.


```vb
Sub AddXmlMap() 
 Dim strSchemaLocation As String 
 
 strSchemaLocation = "http://example.microsoft.com/schemas/CustomerData.xsd" 
 ActiveWorkbook.XmlMaps.Add strSchemaLocation, "Root" 
End Sub
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


