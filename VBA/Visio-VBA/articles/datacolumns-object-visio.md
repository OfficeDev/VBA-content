---
title: DataColumns Object (Visio)
keywords: vis_sdr.chm61015
f1_keywords:
- vis_sdr.chm61015
ms.prod: visio
api_name:
- Visio.DataColumns
ms.assetid: 620a56f5-d552-1247-22fb-18d07993d5ad
ms.date: 06/08/2017
---


# DataColumns Object (Visio)

The collection of  **DataColumn** objects associated with a **DataRecordset** object.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Remarks

The default property of the  **DataColumns** collection is **Item** .

A  **DataRecordset** object can contain only one **DataColumns** collection. The number of **DataColumn** objects that can belong to a **DataColumns** collection is limited only by the number of columns in the data source and the hardware constraints of your computer.

You can use the  **[SetColumnProperties](datacolumns-setcolumnproperties-method-visio.md)** method to set multiple properties of the data recordset columns you specify to the values you specify. Note that **SetColumnProperties** can set values of multiple properties for multiple columns, whereas the **[DataColumn.SetProperty](datacolumn-setproperty-method-visio.md)** method sets the value of only one property of one column at a time.


