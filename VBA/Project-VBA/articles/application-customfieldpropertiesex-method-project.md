---
title: Application.CustomFieldPropertiesEx Method (Project)
keywords: vbapj.chm35
f1_keywords:
- vbapj.chm35
ms.prod: project-server
api_name:
- Project.Application.CustomFieldPropertiesEx
ms.assetid: 3eac9820-848a-011a-96df-f752ea33f31f
ms.date: 06/08/2017
---


# Application.CustomFieldPropertiesEx Method (Project)

Sets attributes associated with a custom field.


## Syntax

 _expression_. **CustomFieldPropertiesEx**( ** _FieldID_**, ** _Attribute_**, ** _SummaryCalc_**, ** _GraphicalIndicators_**, ** _Required_**, ** _AutomaticallyRolldownToAssn_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**Long**|The custom field. Can be one of the  **[PjCustomField](pjcustomfield-enumeration-project.md)** constants.|
| _Attribute_|Optional|**Long**|The attribute to associate with the field specified with FieldID. Can be one of the following  **[PjCustomFieldAttribute](pjcustomfieldattribute-enumeration-project.md)** constants: **pjFieldAttributeNone**, **pjFieldAttributeFormula**, or **pjFieldAttributeValueList**.|
| _SummaryCalc_|Optional|**Long**|The calculation to be performed on the custom field for summary rows and grouping summary rows. Can be one of the  **[PjSummaryCalc](pjsummarycalc-enumeration-project.md)** constants.|
| _GraphicalIndicators_|Optional|**Boolean**|**True** if graphical indicators display instead of data for the custom field.|
| _Required_|Optional|**Boolean**|**True** if the custom field is required.|
| _AutomaticallyRolldownToAssn_|Optional|**Boolean**|True if the custom field automatically rolls down to assignments.|

### Return Value

 **Boolean**


## Remarks

Changing the value of Attribute for a field only enables or disables the attribute. It does not remove any associated data.


