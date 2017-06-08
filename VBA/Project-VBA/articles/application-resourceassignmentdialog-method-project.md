---
title: Application.ResourceAssignmentDialog Method (Project)
keywords: vbapj.chm943
f1_keywords:
- vbapj.chm943
ms.prod: project-server
api_name:
- Project.Application.ResourceAssignmentDialog
ms.assetid: efe91944-bdfa-a15c-6f28-44fe4d629974
ms.date: 06/08/2017
---


# Application.ResourceAssignmentDialog Method (Project)

Displays the  **Assign Resources** dialog box, expands or collapses the **Resource list options**, and specifies fields and filters.


## Syntax

 _expression_. **ResourceAssignmentDialog**( ** _ShowResourceListOptions_**, ** _ResourceListFields_**, ** _UseNamedFilter_**, ** _FilterName_**, ** _UseAvailableToWorkFilter_**, ** _AvailableToWork_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ShowResourceListOptions_|Optional|**Boolean**|**True** if Project expands the **Resource list options**. The default value is  **False**.|
| _ResourceListFields_|Optional|**Long**|The fields listing data from the active project. Can be one of the following  **PjAssignResourcesListFields** constants: **pjAllColumns** or **pjBasic**. The default value is **pjAllColumns**.|
| _UseNamedFilter_|Optional|**Boolean**|**True** if Project filters resource lists by the filter specified in the FilterName argument.|
| _FilterName_|Optional|**String**|A string representing the name of the resource filter to be applied to the resource list.|
| _UseAvailableToWorkFilter_|Optional|**Boolean**|**True** if Project filters the resource list by a resource's availability to work.|
| _AvailableToWork_|Optional|**Variant**|The number of hours a resource is available to work, without the letter indicating the units.|

### Return Value

 **Boolean**


