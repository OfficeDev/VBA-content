---
title: CubeFields.GetMeasure Method (Excel)
keywords: vbaxl10.chm670078
f1_keywords:
- vbaxl10.chm670078
ms.prod: excel
ms.assetid: 26647294-66df-4691-fa8e-d14cb869145b
ms.date: 06/08/2017
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

|**Important**|
|:-----|  
|<p>Getting a measure by using the  **GetMeasure** function will work for these functions only: **Count**,  **Sum**,  **Average**,  **Max** and **Min**. For example:</p><p>These will work</p><ul><li>```Get CubeField0 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlCount, "NumCarsOwnedCount")```</li><li>```Set CubeField1 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlSum, "NumCarsOwnedSum")```</li><li>```Set CubeField2 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlAverage, "NumCarsOwnedAverage")```</li><li>```Set CubeField4 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlMax, "NumCarsOwnedMax")```</li><li>```Set CubeField5 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlMin, "NumCarsOwnedMin")```</li></ul><p>These will not work</p><ul><li>```Set CubeField3 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlCountNums, "NumCarsOwnedCountNums")</li><li>Set CubeField6 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlProduct, "NumCarsOwnedProduct")```</li><li>```Set CubeField7 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlStDev, "NumCarsOwnedStDev")```</li><li>```Set CubeField8 = modelPivotTable.CubeFields.GetMeasure("[customer].[num_cars_owned]", xlStDevP, "NumCarsOwnedStDevP")```</li></ul>|

 
 

 



 
 
 
 
 



### Return value

 **CUBEFIELD**


## See also


#### Concepts


[CubeFields Object](cubefields-object-excel.md)

