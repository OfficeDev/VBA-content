---
title: Application.ImportCommitment Method (Project)
keywords: vbapj.chm2098
f1_keywords:
- vbapj.chm2098
ms.prod: project-server
api_name:
- Project.Application.ImportCommitment
ms.assetid: ad87bf6a-5409-bd10-b658-b81a3ba501f4
ms.date: 06/08/2017
---


# Application.ImportCommitment Method (Project)

Imports the specified deliverable into a project.


## Syntax

 _expression_. **ImportCommitment**( ** _CommitmentDate_**, ** _CommitmentGuid_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CommitmentDate_|Optional|**Variant**|The commitment date of the deliverable.|
| _CommitmentGuid_|Optional|**Variant**|The class identifier of the deliverable.|

### Return Value

 **Boolean**


## Remarks

In this method, the term  _commitment_ is synonymous with the _deliverable_ feature in Project.


