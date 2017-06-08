---
title: Connections.Add Method (Excel)
keywords: vbaxl10.chm776079
f1_keywords:
- vbaxl10.chm776079
ms.prod: excel
api_name:
- Excel.Connections.Add
ms.assetid: 2dff072d-b250-e052-64d7-f75a4746a23f
ms.date: 06/08/2017
---


# Connections.Add Method (Excel)

Adds a new connction to the workbook.


## Syntax

 _expression_ . **Add**( **_Name_** , **_Description_** , **_ConnectionString_** , **_CommandText_** , **_lCmdtype_** ), **_CreateModelConnection_** , **_ImportRelationships_**

 _expression_ A variable that represents a **Connections** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|Name of the connection.|
| _Description_|Required| **String**|Brief description about the connection.|
| _ConnectionString_|Required| **Variant**|The connection string.|
| _CommandText_|Required| **Variant**|The command text to create the connection.|
| _lCmdtype_|Optional| **Variant**|Command type.|
| _CreateModelConnection_|Optional| **Boolean**|Specifies whether to create a connection to the PowerPivot model.|
| _ImportRelationships]_|Optional| **Boolean**|Specifies whether to import any existing relationships.|

### Return Value

WorkbookConnection


## See also


#### Concepts


[Connections Object](connections-object-excel.md)

