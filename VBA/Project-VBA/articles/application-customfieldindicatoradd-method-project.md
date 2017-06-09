---
title: Application.CustomFieldIndicatorAdd Method (Project)
keywords: vbapj.chm38
f1_keywords:
- vbapj.chm38
ms.prod: project-server
api_name:
- Project.Application.CustomFieldIndicatorAdd
ms.assetid: dc5d071b-3cf8-fe56-df16-c5a6051142da
ms.date: 06/08/2017
---


# Application.CustomFieldIndicatorAdd Method (Project)

Creates a test condition against the value of a custom field to determine which graphical indicator to display in place of the actual data.


## Syntax

 _expression_. **CustomFieldIndicatorAdd**( ** _FieldID_**, ** _Test_**, ** _Value_**, ** _IndicatorID_**, ** _CriteriaList_**, ** _Index_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**Long**|The custom field. Can be one of the  **[PjCustomField](pjcustomfield-enumeration-project.md)** constants.|
| _Test_|Required|**Long**|The type of comparison to perform between  **Value** and the custom field specified by **FieldID**. Can be one of the **[PjComparison](pjcomparison-enumeration-project.md)** constants.|
| _Value_|Required|**String**|The value to compare with the custom field's value. If  **Test** is **pjCompareAnyValue**, **Value** is ignored.|
| _IndicatorID_|Required|**Long**|The indicator image to display when the value of the field specified with  **FieldID** passes the comparison specified with **Test**. Can be one of the **[PjIndicator](pjindicator-enumeration-project.md)** constants.|
| _CriteriaList_|Optional|**Long**|The criteria list to which the test condition should be added. Can be one of the  **[PjCriteriaList](pjcriterialist-enumeration-project.md)** constants. The default value is **pjCriteriaNonSummary**.|
| _Index_|Optional|**Integer**|The position to add the test condition in the list specified by  **CriteriaList**. Tests are evaluated in ascending **Index** order. If **Index** is n + 2 or greater, where n is the number of existing tests, the test is added at n + 1. The default value is n + 1.|

### Return Value

 **Boolean**


## Remarks

The  **CustomFieldIndicatorAdd** method returns a trappable error (error code 1004) if the list specified by _CriteriaList_ is read-only because it has been set to inherit values from another list.


