---
title: Application.SelectTPLineHeight Method (Project)
keywords: vbapj.chm70
f1_keywords:
- vbapj.chm70
ms.prod: project-server
api_name:
- Project.Application.SelectTPLineHeight
ms.assetid: f637032a-ede4-6164-e796-716bf5f556f1
ms.date: 06/08/2017
---


# Application.SelectTPLineHeight Method (Project)

Sets the number of text lines for assignment rows in the Team Planner.


## Syntax

 _expression_. **SelectTPLineHeight**( ** _LineMultiple_** )

 _expression_ An expression that returns a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _LineMultiple_|Required|**Integer**|Number of text lines.|

### Return Value

 **Boolean**


## Remarks

The  **Text Lines** drop-down list values range from 1 to 10 in the Team Planner view. However, the _LineMultiple_ argument in the **SelectTPLineHeight** method can range from 1 to 32767.


