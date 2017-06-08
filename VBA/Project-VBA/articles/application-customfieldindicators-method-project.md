---
title: Application.CustomFieldIndicators Method (Project)
keywords: vbapj.chm37
f1_keywords:
- vbapj.chm37
ms.prod: project-server
api_name:
- Project.Application.CustomFieldIndicators
ms.assetid: afbb7bff-49fe-7e12-a257-cab4c730bfbb
ms.date: 06/08/2017
---


# Application.CustomFieldIndicators Method (Project)

Sets options for graphical indicators for a custom field.


## Syntax

 _expression_. **CustomFieldIndicators**( ** _FieldID_**, ** _SummaryInheritsNonsummary_**, ** _ProjectInheritsSummary_**, ** _ShowToolTips_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**Long**|The custom field. Can be one of the  **[PjCustomField](pjcustomfield-enumeration-project.md)** constants.|
| _SummaryInheritsNonsummary_|Optional|**Boolean**|**True** if summary rows use the same test criteria for displaying graphical indicators and use the same images as nonsummary rows. **False** if graphical indicators for summary rows are based on a different set of comparison tests and values and use different images than nonsummary rows.|
| _ProjectInheritsSummary_|Optional|**Boolean**|**True** if the project summary row uses the same test criteria for displaying graphical indicators and uses the same images as summary rows. **False** if graphical indicators for the project summary row are based on a different set of comparison tests and values and use different images than other summary rows.|
| _ShowToolTips_|Optional|**Boolean**|**True** if pausing the mouse over a graphical indicator in a custom field displays a tooltip with the actual data for the custom field.|

### Return Value

 **Boolean**


