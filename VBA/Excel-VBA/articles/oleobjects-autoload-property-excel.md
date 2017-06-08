---
title: OLEObjects.AutoLoad Property (Excel)
keywords: vbaxl10.chm421073
f1_keywords:
- vbaxl10.chm421073
ms.prod: excel
api_name:
- Excel.OLEObjects.AutoLoad
ms.assetid: 0b833fe9-33c6-e97d-3b19-52429ed88d88
ms.date: 06/08/2017
---


# OLEObjects.AutoLoad Property (Excel)

 **True** if the OLE object is automatically loaded when the workbook that contains it is opened. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoLoad**

 _expression_ A variable that represents an **OLEObjects** object.


## Remarks

This property is ignored by ActiveX controls. ActiveX controls are always loaded when a workbook is opened.

For most OLE object types, this property shouldn't be set to  **True** . By default, the **AutoLoad** property is set to **False** for new OLE objects; this saves time and memory when Microsoft Excel is loading workbooks. The benefit of automatically loading OLE objects is that, for objects that represent volatile data, links to source data can be reestablished immediately and the objects can be rendered again, if necessary.


## Example

This example sets the  **AutoLoad** property for OLE object one on the active sheet.


```vb
ActiveSheet.OLEObjects(1).AutoLoad = False
```


## See also


#### Concepts


[OLEObjects Object](oleobjects-object-excel.md)

