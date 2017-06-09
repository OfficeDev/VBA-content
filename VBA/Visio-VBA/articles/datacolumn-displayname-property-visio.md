---
title: DataColumn.DisplayName Property (Visio)
keywords: vis_sdr.chm16760395
f1_keywords:
- vis_sdr.chm16760395
ms.prod: visio
api_name:
- Visio.DataColumn.DisplayName
ms.assetid: eddfba36-836b-97c4-2b34-a5a930d85d03
ms.date: 06/08/2017
---


# DataColumn.DisplayName Property (Visio)

Specifies the name that appears for the data column on the tab of the parent data recordset in the  **External Data** window in the Microsoft Visio user interface. Read/write.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **DisplayName**

 _expression_ An expression that returns a **DataColumn** object.


### Return Value

String


## Remarks

The  **DisplayName** property value is not necessarily the same as that of the **[Name](datacolumn-name-property-visio.md)** property, which is read-only, and which determines the unique identifier of the data column in its parent data recordset.

The value of the  **DisplayName** property corresponds to the value in the Label column in the Shape Data section of the ShapeSheet spreadsheet of a shape that is linked to data. The Label column value determines the label that appears for a particular shape-data item in the **Shape Data** dialog box for that shape. (Right-click the shape, and then click **Shape Data**.)


 **Note**  In some previous versions of Visio, Shape Data were called Custom Properties.


