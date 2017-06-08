---
title: Filters.Copy Method (Project)
keywords: vbapj.chm132248
f1_keywords:
- vbapj.chm132248
ms.prod: project-server
api_name:
- Project.Filters.Copy
ms.assetid: e0432403-a31f-f60a-1a60-c7731809d626
ms.date: 06/08/2017
---


# Filters.Copy Method (Project)

Makes a copy of a group definition for the  **Filters** collection and returns a reference to the **[Filter](filter-object-project.md)** object.


## Syntax

 _expression_. **Copy**( ** _Source_**, ** _NewName_** )

 _expression_ A variable that represents a **Filters** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Source_|Required|**String**|The name of the filter to copy.|
| _NewName_|Required|**String**|The name of the new filter.|

### Return Value

 **Filter**


## See also


#### Concepts


[Filters Collection Object](filters-object-project.md)
