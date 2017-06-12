---
title: Workbook.AutoUpdateSaveChanges Property (Excel)
keywords: vbaxl10.chm199079
f1_keywords:
- vbaxl10.chm199079
ms.prod: excel
api_name:
- Excel.Workbook.AutoUpdateSaveChanges
ms.assetid: 06f9951d-a17a-bf88-4f6e-65835eb112f8
ms.date: 06/08/2017
---


# Workbook.AutoUpdateSaveChanges Property (Excel)

 **True** if current changes to the shared workbook are posted to other users whenever the workbook is automatically updated. **False** if changes aren't posted (this workbook is still synchronized with changes made by other users). The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **AutoUpdateSaveChanges**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

The  **[AutoUpdateFrequency](workbook-autoupdatefrequency-property-excel.md)** property must be set to a value from 5 to 1440 for this property to take effect.


## Example

This example causes changes to the shared workbook to be posted to other users whenever the workbook is automatically updated.


```vb
ActiveWorkbook.AutoUpdateSaveChanges = True
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

