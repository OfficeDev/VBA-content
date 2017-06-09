---
title: Application.CustomFieldValueListAdd Method (Project)
keywords: vbapj.chm41
f1_keywords:
- vbapj.chm41
ms.prod: project-server
api_name:
- Project.Application.CustomFieldValueListAdd
ms.assetid: 6ef6c528-dc7a-00e8-a770-70b3b9ab86ae
ms.date: 06/08/2017
---


# Application.CustomFieldValueListAdd Method (Project)

Adds an item to the value list for a custom field.


## Syntax

 _expression_. **CustomFieldValueListAdd**( ** _FieldID_**, ** _Value_**, ** _Description_**, ** _Phonetic_**, ** _Index_**, ** _FieldDefault_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**Long**|The custom field. Can be one of the  **[PjCustomField](pjcustomfield-enumeration-project.md)** constants.|
| _Value_|Optional|**String**|The value to add to the list.|
| _Description_|Optional|**String**|A description of the value.|
| _Phonetic_|Optional|**String**|The phonetic spelling of  **Value**, used for sorting into syllabary order in Japanese. For languages other than Japanese, **Phonetic** is ignored.|
| _Index_|Optional|**Integer**|The position to add the item specified with  **Value**, relative to other items in the value list. If **Index** is n + 2 or greater, where n is the number of existing items, the item is added at n + 1. The default value is n + 1.|
| _FieldDefault_|Optional|**Boolean**|**True** if the value specified with **Value** functions as the default for the custom field. The default value is **False**.|

### Return Value

 **Boolean**


