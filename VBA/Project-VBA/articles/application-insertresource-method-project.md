---
title: Application.InsertResource Method (Project)
keywords: vbapj.chm2179
f1_keywords:
- vbapj.chm2179
ms.prod: project-server
api_name:
- Project.Application.InsertResource
ms.assetid: e3e62534-3a78-28a2-fb87-ed017b83f9fb
ms.date: 06/08/2017
---


# Application.InsertResource Method (Project)

Inserts a new resource in a resource view.


## Syntax

 _expression_. **InsertResource**( ** _Type_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional|**PjResourceTypes**|Specifies whether the type is a cost, material, or work resource. Can be one of the  **[PjResourceTypes](pjresourcetypes-enumeration-project.md)** constants. The default is **pjResourceTypeWork**.|

### Return Value

 **Boolean**


## Remarks

The  **InsertResource** method corresponds to the **Insert Resource** command on the right-click option menu in the Resource Sheet view or Resource Usage view. The **Resource Name** cell is selected with **<Type Resource Name Here>**. In the Team Planner view,  **InsertResource** creates a row below the last resource, with the name **New Resource**. 


