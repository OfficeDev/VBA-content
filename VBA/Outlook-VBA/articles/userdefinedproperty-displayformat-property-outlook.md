---
title: UserDefinedProperty.DisplayFormat Property (Outlook)
keywords: vbaol11.chm8
f1_keywords:
- vbaol11.chm8
ms.prod: outlook
api_name:
- Outlook.UserDefinedProperty.DisplayFormat
ms.assetid: f891aa8d-a769-275d-c027-7c5260eafc97
ms.date: 06/08/2017
---


# UserDefinedProperty.DisplayFormat Property (Outlook)

Returns a  **Long** value that represents the display format for the **[UserDefinedProperty](userdefinedproperty-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **DisplayFormat**

 _expression_ A variable that represents a **UserDefinedProperty** object.


## Remarks

The value of this property is a constant from an enumeration, where the enumeration is dependent on the value of the  **[Type](userdefinedproperty-type-property-outlook.md)** property for the **UserDefinedProperty** object:



| **Type value**| **DisplayFormat enumeration**|
| **olCombination**|No enumeration available. This property always returns 1 for  **olCombination** .|
| **olCurrency**| **[OlFormatCurrency](olformatcurrency-enumeration-outlook.md)**|
| **olDateTime**| **[OlFormatDateTime](olformatdatetime-enumeration-outlook.md)**|
| **olDuration**| **[OlFormatDuration](olformatduration-enumeration-outlook.md)**|
| **olEnumeration**| **[OlFormatEnumeration](olformatenumeration-enumeration-outlook.md)**|
| **olFormula**|No enumeration available. This property always returns 1 for  **olFormula** .|
| **olInteger**| **[OlFormatInteger](olformatinteger-enumeration-outlook.md)**|
| **olKeywords**| **[OlFormatKeywords](olformatkeywords-enumeration-outlook.md)**|
| **olNumber**| **[OlFormatNumber](olformatnumber-enumeration-outlook.md)**|
| **olOutlookInternal**|No enumeration available. This property always returns 1 for  **olOutlookInternal** .|
| **olPercent**| **[OlFormatPercent](olformatpercent-enumeration-outlook.md)**|
| **olSmartFrom**| **[OlFormatSmartFrom](olformatsmartfrom-enumeration-outlook.md)**|
| **olText**| **[OlFormatText](olformattext-enumeration-outlook.md)**|
| **olYesNo**| **[OlFormatYesNo](olformatyesno-enumeration-outlook.md)**|

## See also


#### Concepts


[UserDefinedProperty Object](userdefinedproperty-object-outlook.md)

