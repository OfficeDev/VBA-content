---
title: Project.DeliverableUpdate Method (Project)
ms.prod: project-server
api_name:
- Project.Project.DeliverableUpdate
ms.assetid: 665e79a0-b3b4-e36e-6369-627e526f7db0
ms.date: 06/08/2017
---


# Project.DeliverableUpdate Method (Project)

Updates the properties of a deliverable.


## Syntax

 _expression_. **DeliverableUpdate**( ** _DeliverableGuid_**, ** _DeliverableName_**, ** _DeliverableStartDate_**, ** _DeliverableFinishDate_** )

 _expression_ A variable that represents a **Project** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DeliverableGuid_|Required|**String**|GUID of the deliberable to update.|
| _DeliverableName_|Required|**String**|Name of the deliverable.|
| _DeliverableStartDate_|Required|**Variant**|Date when the deliverable starts.|
| _DeliverableFinishDate_|Required|**Variant**|Date when the deliverable is finished.|

### Return Value

 **Boolean**


