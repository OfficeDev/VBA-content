---
title: XmlMap.Delete Method (Excel)
keywords: vbaxl10.chm754086
f1_keywords:
- vbaxl10.chm754086
ms.prod: excel
api_name:
- Excel.XmlMap.Delete
ms.assetid: 8acde534-c465-029a-635a-38f63c5f4013
ms.date: 06/08/2017
---


# XmlMap.Delete Method (Excel)

Removes the specified XML map from the workbook.


## Syntax

 _expression_ . **Delete**

 _expression_ A variable that represents a **XmlMap** object.


## Remarks

Deleting the XML map will convert all of the XML Lists to generic Lists and remove all of the single-cell mappings (with the data still remaining). In addition, the  **[XmlMap](xmlmap-object-excel.md)** object will be removed from the **[XmlMaps](xmlmaps-object-excel.md)** collection. The map and schema information will be removed from the workbook (it will no longer be persisted in the XLS file and XMLSS). Any references to the deleted object become invalid.


## See also


#### Concepts


[XmlMap Object](xmlmap-object-excel.md)

