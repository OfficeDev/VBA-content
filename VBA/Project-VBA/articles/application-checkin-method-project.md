---
title: Application.CheckIn Method (Project)
keywords: vbapj.chm2323
f1_keywords:
- vbapj.chm2323
ms.prod: project-server
api_name:
- Project.Application.CheckIn
ms.assetid: dd2cc86f-44f5-9c7e-c4d1-8475d11367ac
ms.date: 06/08/2017
---


# Application.CheckIn Method (Project)

Checks in the active project file if it is stored in a SharePoint library.


## Syntax

 _expression_. **CheckIn**( ** _fSaveChanges_**, ** _Comments_**, ** _fMakePublic_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _fSaveChanges_|Optional|**Boolean**|**True** saves changes and checks in the project document. **False** returns the document to a checked-in status without saving revision.|
| _Comments_|Optional|**String**| Allows the user to enter check-in comments for the revision of the project being checked in (applies only if fSaveChanges equals **True** ).|
| _fMakePublic_|Optional|**Boolean**|**True** allows the user to publish the project after it has been checked in. This submits the project for the approval process, which can eventually result in a version of the project being published to users with read-only rights to the project (applies only if fSaveChanges equals **True** ).|

### Return Value

 **Boolean**


