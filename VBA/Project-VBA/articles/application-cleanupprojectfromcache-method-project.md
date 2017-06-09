---
title: Application.CleanupProjectFromCache Method (Project)
keywords: vbapj.chm2191
f1_keywords:
- vbapj.chm2191
ms.prod: project-server
api_name:
- Project.Application.CleanupProjectFromCache
ms.assetid: 40fef64a-036f-8e1c-ce86-0c3609777f77
ms.date: 06/08/2017
---


# Application.CleanupProjectFromCache Method (Project)

Deletes the specified project file from the local cache. Available only in Project Professional.


## Syntax

 _expression_. **CleanupProjectFromCache**( ** _Filename_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Optional|**String**|Name of the project file to delete from the cache.|

### Return Value

Boolean


## Remarks

You can use the  **CleanupProjectFromCache** method if you suspect the project in the local cache is corrupted. If _FileName_ is omitted, **CleanupProjectFromCache** does nothing.


## Example




```
CleanupProjectFromCache("Sample.mpp")
```


