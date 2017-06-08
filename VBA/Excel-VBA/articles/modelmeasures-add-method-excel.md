---
title: ModelMeasures.Add Method (Excel)
keywords: vbaxl10.chm980077
f1_keywords:
- vbaxl10.chm980077
ms.assetid: abc0f260-abdb-2f60-928f-b325fbb976f3
ms.date: 06/08/2017
ms.prod: excel
---


# ModelMeasures.Add Method (Excel)

Adds a model measure to the model.


## Syntax

 _expression_ . **Add**( _MeasureName_,  _MeasureName_,  _AssociatedTable_,  _Formula_,  _FormatInformation_,  _Description_)

 _expression_ A variable that represents a **ModelMeasures** object.


### Parameters



|||||
| _MeasureName_|Required|STRING|The name of the model measure.|
| _AssociatedTable_|Required|MODELTABLE|The model table associated with the model measure. This is the table that contains the model measure, as seen in the  **Field List** task pane.|
| _Formula_|Required|STRING|The Data Analysis Expressions (DAX) formula, inserted as a string.|
| _FormatInformation_|Required|VARIANT|The formatting of the model measure. See Remarks. |
| _Description_|Optional|VARIANT|The description associated with the model measure.|

### Return Value

[ModelMeasure](modelmeasure-object-excel.md)


## Remarks

The formatting of the model measure can be of type [ModelFormatBoolean](modelformatboolean-object-excel.md), [ModelFormatCurrency](modelformatcurrency-object-excel.md), [ModelFormatDate](modelformatdate-object-excel.md), [ModelFormatDecimalNumber](modelformatdecimalnumber-object-excel.md), [ModelFormatGeneral](modelformatgeneral-object-excel.md), [ModelFormatPercentageNumber](modelformatpercentagenumber-object-excel.md), [ModelFormatScientificNumber](modelformatscientificnumber-object-excel.md), or [ModelFormatWholeNumber](modelformatwholenumber-object-excel.md).


## See also


#### Other resources


[ModelMeasures Object ](modelmeasures-object-excel.md)


