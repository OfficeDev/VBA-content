---
title: Page.SplitConnector Method (Visio)
keywords: vis_sdr.chm10962155
f1_keywords:
- vis_sdr.chm10962155
ms.prod: visio
api_name:
- Visio.Page.SplitConnector
ms.assetid: b2d371b5-3769-00cd-688f-2391a8c504e9
ms.date: 06/08/2017
---


# Page.SplitConnector Method (Visio)

Splits the specified connector with the specified shape. Returns the new duplicated connector.


## Syntax

 _expression_ . **SplitConnector**( **_ConnectorToSplit_** , **_Shape_** )

 _expression_ A variable that represents a **[Page](page-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ConnectorToSplit_|Required| **[Shape](shape-object-visio.md)**|The connector to split. Must be a routable one-dimensional (1-D) connector.|
| _Shape_|Required| **Shape**|The shape to use to split the connector. Must be a two-dimensional (2-D) shape.|

### Return Value

 **Shape**


## Remarks

When you call  **SplitConnector** , Microsoft Visio duplicates the connector (except for text), positions the shape automatically, and glues the shape to the two resulting connectors.


