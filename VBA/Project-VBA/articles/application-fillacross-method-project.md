---
title: Application.FillAcross Method (Project)
keywords: vbapj.chm244
f1_keywords:
- vbapj.chm244
ms.prod: project-server
api_name:
- Project.Application.FillAcross
ms.assetid: 9ab6a32a-84b4-e9c5-2632-b02205275e82
ms.date: 06/08/2017
---


# Application.FillAcross Method (Project)

Fills the selected cells or columns with the values in the specified cell or column of the selection.


## Syntax

 _expression_. **FillAcross**( ** _Right_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Right_|Optional|**Boolean**|**True** if values in the leftmost cell or column of the selection are copied right to the other selected cells or columns. **False** if values in the rightmost cell or column are copied left to the other selected cells or columns. The default value is **True**.|

### Return Value

 **Boolean**


## Remarks

The  **FillAcross** method is only available in timephased cells of usage views.


