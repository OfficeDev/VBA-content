---
title: Project.CheckIn Method (Project)
keywords: vbapj.chm132592
f1_keywords:
- vbapj.chm132592
ms.prod: project-server
api_name:
- Project.Project.CheckIn
ms.assetid: 9620bd94-4b75-5c7e-2993-5018c5bb84e3
ms.date: 06/08/2017
---


# Project.CheckIn Method (Project)

Checks in the working copy of the project from a local computer to the SharePoint document library, and sets the local project to read-only so that it cannot be edited locally.


## Syntax

 _expression_. **CheckIn**( ** _SaveChanges_**, ** _Comment_**, ** _MakePublic_** )

 _expression_ A variable that represents a **Project** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SaveChanges_|Optional|**Boolean**|**True** saves changes and checks in the document. **False** returns the document to a checked-in status without saving revisions.|
| _Comment_|Optional|**String**|Comments for the revision of the project being checked in (applies only if SaveChanges equals  **True** ).|
| _MakePublic_|Optional|**Boolean**|**True** allows the user to publish the project after it has been checked in. This submits the project for the approval process, which can eventually result in a version of the project being published to users with read-only rights to the project (applies only if SaveChanges equals **True** ).|

## Remarks

 The **CheckIn** method also closes the project after it is checked in.


