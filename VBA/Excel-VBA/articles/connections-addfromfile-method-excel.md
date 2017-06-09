---
title: Connections.AddFromFile Method (Excel)
keywords: vbaxl10.chm776080
f1_keywords:
- vbaxl10.chm776080
ms.prod: excel
api_name:
- Excel.Connections.AddFromFile
ms.assetid: 24b13d14-6e5e-ee76-d4a9-79441647a803
ms.date: 06/08/2017
---


# Connections.AddFromFile Method (Excel)

Adds a connection from the specified file.


## Syntax

 _expression_ . **AddFromFile**( **_Filename_** , **_Filename_** , **_Filename_** )

 _expression_ A variable that represents a **Connections** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Required| **String**|Name of the file.|
| _CreateModelConnection_|Optional| **Boolean**|Specifies whether to create the connection to the model.|
| _ImportRelationships_|Optional| **Boolean**|Specifies whether to import the connection relationship.|

### Return Value

WorkbookConnection


## See also


#### Concepts


[Connections Object](connections-object-excel.md)

