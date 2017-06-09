---
title: ReportTable.UpdateTableData Method (Project)
keywords: vbapj.chm132700
f1_keywords:
- vbapj.chm132700
ms.prod: project-server
ms.assetid: 5a5b1ed3-779e-7be5-6bd5-2ba544e0d27f
ms.date: 06/08/2017
---


# ReportTable.UpdateTableData Method (Project)
Updates rows and columns in the report table to show the specified task or resource fields; uses the specified filter, group, and outline level.

## Syntax

 _expression_. **UpdateTableData** _(Task,_ _GroupName,_ _FilterName,_ _OutlineLevel,_ _SafeArrayOfPjField)_

 _expression_ A variable that represents a **ReportTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Task_|Required|**Boolean**|**True** to update task data in the **Field List** task pane; **False** to update resource data.|
| _GroupName_|Optional|**String**|Name of the group in the  **Group By** drop-down list.|
| _FilterName_|Optional|**String**|Name of the filter in the  **Filter** drop-down list.|
| _OutlineLevel_|Optional|**Long**|The task outline level. The default value is -1, which the equivalent of  **Show All**. Not used for resource fields (when  _Task_ is **False**).|
| _SafeArrayOfPjField_|Optional|**Variant**|Specifies an array of fields for the update, where each item in the array can be a [PjField](pjfield-enumeration-project.md) constant.|
| _Task_|Required|BOOL||
| _GroupName_|Optional|STRING||
| _FilterName_|Optional|STRING||
| _OutlineLevel_|Optional|INT||
| _SafeArrayOfPjField_|Optional|VARIANT||

### Return value

 **Nothing**


## See also


#### Other resources


[ReportTable Object](reporttable-object-project.md)
