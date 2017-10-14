---
title: Application.JobStart Event (Project)
ms.prod: project-server
api_name:
- Project.Application.JobStart
ms.assetid: 874b35cb-bb90-b8dc-3c22-84c8809c3177
ms.date: 06/08/2017
---


# Application.JobStart Event (Project)

Occurs before the queue job is put on the server queue. Project Professional only.


## Syntax

 _expression_. **JobStart**( ** _bstrName_**, ** _bstrprojGuid_**, ** _bstrjobGuid_**, ** _jobType_**, ** _lResult_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _bstrName_|Required|**String**|Name of the project whose queue job was completed.|
| _bstrprojGuid_|Required|**String**|GUID of the project whose queue job was completed.|
| _bstrjobGuid_|Required|**String**|GUID of the job that was completed.|
| _jobType_|Required|**Long**|Job Type of the job that was completed. For example,  **Project Save**, **Project Publish**, **Project Status Update**.|
| _lResult_|Required|**Long**|**HResult** (error code) of the queue operation. For example, **0** indicates that the job succeeded, **E_FAIL** indicates failure|

### Return Value

nothing


## Remarks

More details about the Queue job can be obtained by making the ** getJobCompletionState PSI** call on the **QueueSystem.asmx** webservice with the job GUID.


