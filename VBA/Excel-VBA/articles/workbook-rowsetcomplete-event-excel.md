---
title: Workbook.RowsetComplete Event (Excel)
keywords: vbaxl10.chm503101
f1_keywords:
- vbaxl10.chm503101
ms.prod: excel
api_name:
- Excel.Workbook.RowsetComplete
ms.assetid: 05bdddba-6716-4bba-01b6-863f27623821
ms.date: 06/08/2017
---


# Workbook.RowsetComplete Event (Excel)

The event is raised when the user either drills through the recordset or invokes the rowset action on an OLAP PivotTable.


## Syntax

 _expression_ . **RowsetComplete**( **_Description_** , **_Sheet_** , **_Success_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Description_|Required| **String**|A brief description of the event.|
| _Sheet_|Required| **String**|Worksheet on which the recordset is created.|
| _Success_|Required| **Boolean**|Contains a Boolean value to indicate success or failure.|

## Remarks

Because the recordset is created asynchronously, the event allows automation to determine when the action has been completed. Additionally, because the recordset is created on a separate sheet, the event needs to be on the workbook level.


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

