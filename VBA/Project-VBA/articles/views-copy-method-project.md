---
title: Views.Copy Method (Project)
ms.prod: project-server
api_name:
- Project.Views.Copy
ms.assetid: 5e82641a-f5c6-41a6-23bf-61220a4fc30c
ms.date: 06/08/2017
---


# Views.Copy Method (Project)

Makes a copy of a group definition for the  **Views** collection and returns a reference to the **[View](view-object-project.md)** object.


## Syntax

 _expression_. **Copy**( ** _Source_**, ** _NewName_** )

 _expression_ A variable that represents a **Views** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Source_|Required|**String**|The name of the view to copy.|
| _NewName_|Required|**String**|The name of the new view.|

### Return Value

 **View**


## See also


#### Concepts


[Views Collection Object](views-object-project.md)
