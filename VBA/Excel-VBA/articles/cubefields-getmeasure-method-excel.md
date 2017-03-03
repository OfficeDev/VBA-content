---
title: CubeFields.GetMeasure Method (Excel)
keywords: vbaxl10.chm670078
f1_keywords:
- vbaxl10.chm670078
ms.prod: EXCEL
ms.assetid: 26647294-66df-4691-fa8e-d14cb869145b
---


# CubeFields.GetMeasure Method (Excel)

Given an attribute hierarchy, returns an implicit measure for the given function that corresponds to this attribute. If an ?implicit measure? does not exist, a new implicit measure is created and added to the [CubeFields Object (Excel)](cubefields-object-excel.md) collection.


## Syntax

 _expression_ . **GetMeasure**_(AttributeHierarchy,_ _Function,_ _Caption)_

 _expression_ A variable that represents a **CubeFields** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _AttributeHierarchy_|Required|VARIANT|The unique cube field that is an attribute hierarchy (XlCubeFieldType = xlHierarchy and XlCubeFieldSubType = xlCubeAttribute).|
| _Function_|Required|XLCONSOLIDATIONFUNCTION|The function performed in the added data field.|
| _Caption_|Optional|VARIANT|The label used in the PivotTable report to identify this measure. If the measure already exists, caption will overwrite the existing label of this measure.|

### Remarks


 **Important**  Getting a measure by using the  **GetMeasure** function will work for these functions only: **Count**,  **Sum**,  **Average**,  **Max** and **Min**. For example:These will workThese will not work


### Return value

 **CUBEFIELD**


## See also


#### Concepts


[CubeFields Object](cubefields-object-excel.md)

