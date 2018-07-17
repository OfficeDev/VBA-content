---
title: DataRecordset.DataColumns Property (Visio)
keywords: vis_sdr.chm16460285
f1_keywords:
- vis_sdr.chm16460285
ms.prod: visio
api_name:
- Visio.DataRecordset.DataColumns
ms.assetid: d22c07b9-3c92-fed4-72ed-6676ea64f1bf
ms.date: 06/08/2017
---


# DataRecordset.DataColumns Property (Visio)

Returns the  **[DataColumns](datacolumns-object-visio.md)** collection associated with the **DataRecordset** object. Read-only.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **DataColumns**

 _expression_ An expression that returns a **DataRecordset** object.


### Return Value

DataColumns


## Remarks

Every  **DataRecordset** object contains a **DataColumns** collection of all the **[DataColumn](datacolumn-object-visio.md)** objects associated with the **DataRecordset** object. These objects allow you to map data columns to cells in the Shape Data (formerly Custom Properties) section of the Visio ShapeSheet spreadsheet.

Once you get the  **DataColumns** collection, you can use its **[SetColumnProperties](datacolumns-setcolumnproperties-method-visio.md)** method to set the properties of multiple data columns, or you can get and set the properties of individual data columns by using the **[DataColumn.GetProperty](datacolumn-getproperty-method-visio.md)** and **[DataColumn.SetProperty](datacolumn-setproperty-method-visio.md)** properties respectively.


