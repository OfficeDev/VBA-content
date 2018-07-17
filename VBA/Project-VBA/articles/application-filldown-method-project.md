---
title: Application.FillDown Method (Project)
keywords: vbapj.chm218
f1_keywords:
- vbapj.chm218
ms.prod: project-server
api_name:
- Project.Application.FillDown
ms.assetid: 5ccb5f67-64c1-9230-ca58-52bd9bd2c4d5
ms.date: 06/08/2017
---


# Application.FillDown Method (Project)

Fills the selected cells or rows with the values in the specified cell or row of the selection.


## Syntax

 _expression_. **FillDown**( ** _Down_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Down_|Optional|**Boolean**|**True** if values in the top cell or row of the selection are copied down to the other selected cells or rows. **False** if values in the bottom cell or row of the selection are copied up to the other selected cells or rows. The default value is **True**.|

### Return Value

 **Boolean**


