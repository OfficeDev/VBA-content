---
title: Application.Publish Method (Project)
keywords: vbapj.chm2278
f1_keywords:
- vbapj.chm2278
ms.prod: project-server
api_name:
- Project.Application.Publish
ms.assetid: 8605f6c9-8710-0c08-79c8-8dec2bedfe18
ms.date: 06/08/2017
---


# Application.Publish Method (Project)

Sends message to the Project Server Queueing Service to publish the local project cache to Project Server, and optionally to the associated SharePoint site for the project workspace. Project Professional only. 


## Syntax

 _expression_. **Publish**( ** _Republish_**, ** _WssUrl_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Republish_|Optional|**Boolean**|If  **True**, publish the entire project plan.|
| _WssUrl_|Optional|**String**|URL for the SharePoint site where the project workspace is to be provisioned. If NULL, no workspace is provisioned.|

### Return Value

 **Boolean**


## Remarks

Typically the publish action is incremental: Only changed data is pushed through from the working store to the published store, Republish forces all data to be published. 


