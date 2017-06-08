---
title: OlUserPropertyType Enumeration (Outlook)
keywords: vbaol11.chm3089
f1_keywords:
- vbaol11.chm3089
ms.prod: outlook
api_name:
- Outlook.OlUserPropertyType
ms.assetid: 24a4517a-3e6c-67be-33a3-fc9c2fb3f1d1
ms.date: 06/08/2017
---


# OlUserPropertyType Enumeration (Outlook)

Indicates the user property type.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olCombination**|19|The property type is a combination of other types. It corresponds to the MAPI type  **PT_STRING8**.|
| **olCurrency**|14|Represents a  **Currency** property type. It corresponds to the MAPI type **PT_CURRENCY**.|
| **olDateTime**|5|Represents a  **DateTime** property type. It corresponds to the MAPI type **PT_SYSTIME**.|
| **olDuration**|7|Represents a time duration property type. It corresponds to the MAPI type  **PT_LONG**.|
| **olEnumeration**|21|Represents an enumeration property type. It corresponds to the MAPI type  **PT_LONG**.|
| **olFormula**|18|Represents a formula property type. It corresponds to the MAPI type  **PT_STRING8**. See  **[UserDefinedProperty.Formula](userdefinedproperty-formula-property-outlook.md)** property.|
| **olInteger**|20|Represents an  **Integer** number property type. It corresponds to the MAPI type **PT_LONG**.|
| **olKeywords**|11|Represents a  **String** array property type used to store keywords. It corresponds to the MAPI type **PT_MV_STRING8**.|
| **olNumber**|3|Represents a  **Double** number property type. It corresponds to the MAPI type **PT_DOUBLE**.|
| **olOutlookInternal**|0|Represents an Outlook internal property type. |
| **olPercent**|12|Represents a  **Double** number property type used to store a percentage. It corresponds to the MAPI type **PT_LONG**.|
| **olSmartFrom**|22|Represents a smart from property type. This property indicates that if the  **From** property of an Outlook item is empty, then the **To** property should be used instead.|
| **olText**|1|Represents a  **String** property type. It corresponds to the MAPI type **PT_STRING8**.|
| **olYesNo**|6|Represents a yes/no ( **Boolean**) property type. It corresponds to the MAPI type  **PT_BOOLEAN**.|

## Remarks

Used by the [ItemProperties.Add](itemproperties-add-method-outlook.md), [UserDefinedProperties.Add](userdefinedproperties-add-method-outlook.md), and [UserProperties.Add](userproperties-add-method-outlook.md) methods, and[ColumnFormat.FieldType](columnformat-fieldtype-property-outlook.md), [ItemProperty.Type](itemproperty-type-property-outlook.md), and [UserDefinedProperty.Type](userdefinedproperty-type-property-outlook.md) properties.


