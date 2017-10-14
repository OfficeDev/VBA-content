---
title: DoCmd.SetProperty Method (Access)
keywords: vbaac10.chm5775
f1_keywords:
- vbaac10.chm5775
ms.prod: access
api_name:
- Access.DoCmd.SetProperty
ms.assetid: 32347eb6-115d-36c5-4c18-eab7e7422b78
ms.date: 06/08/2017
---


# DoCmd.SetProperty Method (Access)

The  **SetProperty** method carries out the SetProperty action in Visual Basic.


## Syntax

 _expression_. **SetProperty**( ** _ControlName_**, ** _Property_**, ** _Value_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ControlName_|Required|**Variant**|The name of the field or control for which you want to set the property value. Leave this argument blank to set the property for the current form or report.|
| _Property_|Optional|**Variant**|A  **[AcProperty](acproperty-enumeration-access.md)** constant that specifies the property that you want to set.|
| _Value_|Optional|**Variant**|The value to which the property is to be set. For properties whose values are either Yes or No, use ?1 for Yes and 0 for No.|

## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

