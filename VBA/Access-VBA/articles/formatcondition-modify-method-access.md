---
title: FormatCondition.Modify Method (Access)
keywords: vbaac10.chm10062
f1_keywords:
- vbaac10.chm10062
ms.prod: access
api_name:
- Access.FormatCondition.Modify
ms.assetid: 213a50f2-30ae-bcdc-d690-2d45bbe6f6e7
ms.date: 06/08/2017
---


# FormatCondition.Modify Method (Access)

You can use the  **Modify** method to change the format conditions of a **[FormatCondition](formatcondition-object-access.md)** object in the **[FormatConditions](formatconditions-object-access.md)** collection of a combo box or text box control.


## Syntax

 _expression_. **Modify**( ** _Type_**, ** _Operator_**, ** _Expression1_**, ** _Expression2_** )

 _expression_ A variable that represents a **FormatCondition** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**AcFormatConditionType**|A  **[AcFormatConditionType](acformatconditiontype-enumeration-access.md)** constant that specifies the type of condition to be modified.|
| _Operator_|Optional|**AcFormatConditionOperator**|A  **[AcFormatConditionOperator](acformatconditionoperator-enumeration-access.md)** constant that specifies the type of operator to be used.
 **Note**  If the type argument is  **acExpression**, the operator argument is ignored. If you leave this argument blank, the default constant ( **acBetween** ) is assumed.

|
| _Expression1_|Optional|**Variant**|A value or expression associated with the first part of the conditional format. Can be a constant value or a string value.|
| _Expression2_|Optional|**Variant**|A value or expression associated with the second part of the conditional format when the operator argument is  **acBetween** or **acNotBetween** (otherwise, this argument is ignored). Can be a constant value or a string value.|

### Return Value

Nothing


## See also


#### Concepts


[FormatCondition Object](formatcondition-object-access.md)

