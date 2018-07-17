---
title: Shapes.AddCatalogMergeFieldToCanvas Method (Publisher)
keywords: vbapb10.chm2162760
f1_keywords:
- vbapb10.chm2162760
ms.prod: publisher
api_name:
- Publisher.Shapes.AddCatalogMergeFieldToCanvas
ms.assetid: 30cd45d0-97f0-ab01-31c2-8d819b435b1b
ms.date: 06/08/2017
---


# Shapes.AddCatalogMergeFieldToCanvas Method (Publisher)

Adds a catalog merge field of the specified type to the canvas. Returns nothing.


## Syntax

 _expression_. **AddCatalogMergeFieldToCanvas**( **_CanvasId_**,  **_CatalogMergeFieldType_**,  **_DbCol_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|CanvasId|Required| **[INT]**|The ID of the canvas to which to add the catalog merge field.|
|CatalogMergeFieldType|Required| **pbCatalogMergeFieldType**|The type (picture or text) of the catalog merge field to add.|
|DbCol|Required| **[INT]**|The number of the column in the data source that contains the catalog merge information.|

