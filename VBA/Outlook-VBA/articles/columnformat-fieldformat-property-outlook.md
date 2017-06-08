---
title: ColumnFormat.FieldFormat Property (Outlook)
keywords: vbaol11.chm2729
f1_keywords:
- vbaol11.chm2729
ms.prod: outlook
api_name:
- Outlook.ColumnFormat.FieldFormat
ms.assetid: 14064b56-65c2-1c7d-1e74-3bfa2d2ccaa7
ms.date: 06/08/2017
---


# ColumnFormat.FieldFormat Property (Outlook)

Returns or sets a  **Long** value that represents the display format of the property to which the **[ColumnFormat](columnformat-object-outlook.md)** object is associated. Read/write.


## Syntax

 _expression_ . **FieldFormat**

 _expression_ A variable that represents a **ColumnFormat** object.


## Remarks

The value of this property is a constant from an enumeration, where the enumeration is dependent on the value of the  **[FieldType](columnformat-fieldtype-property-outlook.md)** property for the **ColumnFormat** object:



| **FieldType value**| **FieldFormat enumeration**|
| **olCurrency**| **[OlFormatCurrency](olformatcurrency-enumeration-outlook.md)**|
| **olFormatDateTime**| **[OlFormatDateTime](olformatdatetime-enumeration-outlook.md)**|
| **olDuration**| **[OlFormatDuration](olformatduration-enumeration-outlook.md)**|
| **olInteger**| **[OlFormatInteger](olformatinteger-enumeration-outlook.md)**|
| **olKeywords**| **[OlFormatKeywords](olformatkeywords-enumeration-outlook.md)**|
| **olNumber**| **[OlFormatNumber](olformatnumber-enumeration-outlook.md)**|
| **olPercent**| **[OlFormatPercent](olformatpercent-enumeration-outlook.md)**|
| **olText**| **[OlFormatText](olformattext-enumeration-outlook.md)**|
| **olYesNo**| **[OlFormatYesNo](olformatyesno-enumeration-outlook.md)**|
| **olEnumeration**| **[OlFormatEnumeration](olformatenumeration-enumeration-outlook.md)**|
| **olSmartFrom**| **[OlFormatSmartFrom](olformatsmartfrom-enumeration-outlook.md)**|
For field types not listed in the above table, the value of this property is ignored.


## See also


#### Concepts


[ColumnFormat Object](columnformat-object-outlook.md)

