---
title: PayRates.Add Method (Project)
ms.prod: project-server
api_name:
- Project.PayRates.Add
ms.assetid: ba5d2667-7452-f9d9-032e-bb7c9d1d4911
ms.date: 06/08/2017
---


# PayRates.Add Method (Project)

Adds a  **PayRate** object to a **PayRates** collection.


## Syntax

 _expression_. **Add**( ** _EffectiveDate_**, ** _StdRate_**, ** _OvtRate_**, ** _CostPerUse_** )

 _expression_ A variable that represents a **PayRates** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EffectiveDate_|Required|**Variant**|The date the new rate comes into effect.|
| _StdRate_|Optional|**Variant**|The new standard rate. If not specified, the rate is the same as for the previous date period.|
| _OvtRate_|Optional|**Variant**|The new overtime rate. If not specified, the rate is the same as for the previous date period.|
| _CostPerUse_|Optional|**Variant**|The new cost per use. If not specified, the rate is the same as for the previous date period.|

### Return Value

 **PayRate**


## See also


#### Concepts


[PayRates Collection Object](payrates-object-project.md)
