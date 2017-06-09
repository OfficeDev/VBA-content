---
title: Application.WebAddToFavorites Method (Project)
keywords: vbapj.chm1314
f1_keywords:
- vbapj.chm1314
ms.prod: project-server
api_name:
- Project.Application.WebAddToFavorites
ms.assetid: 3cf8b3e7-4dbf-8555-1662-2412e7d420b0
ms.date: 06/08/2017
---


# Application.WebAddToFavorites Method (Project)

Adds a link to the current document or selection to the Favorites folder for the user. 


## Syntax

 _expression_. **WebAddToFavorites**( ** _CurrentLink_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CurrentLink_|Optional|**Boolean**|**True** if a link will be added to the current selection. **False** if a link will be added to the current document. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

The Favorites folder is typically  `C:\Users\UserAlias\Favorites`. For a project file named Basic.mpp that is saved in the  `E:\Project\VBA` folder, **WebAddToFavorites** adds a link named Basic that has the following URL: `file:///E:/Project/VBA/Samples/Basic.mpp`

The  **WebAddToFavorites** method is unavailable if the file has never been saved.


