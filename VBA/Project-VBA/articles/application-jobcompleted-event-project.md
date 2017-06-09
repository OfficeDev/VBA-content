---
title: Application.JobCompleted Event (Project)
ms.prod: project-server
api_name:
- Project.Application.JobCompleted
ms.assetid: 44f7987c-92e0-a302-a775-7e62dab2ef86
ms.date: 06/08/2017
---


# Application.JobCompleted Event (Project)

Occurs when a queued job originating from Project Professional is completed.


## Syntax

 _expression_. **JobCompleted**( ** _bstrName_**, ** _bstrprojGuid_**, ** _bstrjobGuid_**, ** _jobType_**, ** _lResult_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _bstrName_|Required|**String**|Name of the project whose queue job was completed.|
| _bstrprojGuid_|Required|**String**|GUID of the project whose queue job was completed.|
| _bstrjobGuid_|Required|**String**|GUID of the job that was completed.|
| _jobType_|Required|**Long**|Type of the job that was completed. For example,  **Project Save**, **Project Publish**, or **Project Status Update**.|
| _lResult_|Required|**Long**|**HResult** (error code) of the queue operation. For example, **0** indicates success and **E_FAIL** indicates failure.|

### Return Value

nothing


