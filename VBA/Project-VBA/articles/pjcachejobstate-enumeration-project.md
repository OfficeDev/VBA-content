---
title: PjCacheJobState Enumeration (Project)
ms.prod: project-server
ms.assetid: 48572c9f-8c3d-8f6d-f633-94f7fedcfe3b
ms.date: 06/08/2017
---


# PjCacheJobState Enumeration (Project)
Contains constants that specify the cache status of a job that Project Professional sends to the Project Server Queue Service.

## Members



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
|**pjCacheJobStateCancelled**|9|The queue job is cancelled.|
|**pjCacheJobStateCorrelationBlocked**|8|The queue job correlation is blocked; the job is not processing.|
|**pjCacheJobStateFailed**|5|The queue job failed.|
|**pjCacheJobStateFailedNotBlocking**|6|The queue job failed but is not blocking continued processing.|
|**pjCacheJobStateInvalid**|-1|The queue job is not valid. (The hexadecimal value is &;HFFFFFFFF.)|
|**pjCacheJobStateLastState**|13|The queue job state is the same as the previous state.|
|**pjCacheJobStateOnHold**|10|The queue job is on hold.|
|**pjCacheJobStateProcessing**|3|The queue job is processing.|
|**pjCacheJobStateReadyForLaunch**|12|The queue job is ready for launch.|
|**pjCacheJobStateReadyForProcessing**|1|The queue job is ready for processing.|
|**pjCacheJobStateSendIncomplete**|2|The queue job is not completely sent to the Project Server Queue Service.|
|**pjCacheJobStateSkipped**|7|The queue job is deferred while another job is being processed.|
|**pjCacheJobStateSleeping**|11|The queue job is sleeping.|
|**pjCacheJobStateSuccess**|4|The queue job completed successfully.|
|**pjCacheJobStateUnknown**|0|The queue job state is unknown.|
|**pjCacheJobStateCancelled**|**9**||
|**pjCacheJobStateCorrelationBlocked**|**8**||
|**pjCacheJobStateFailed**|**5**||
|**pjCacheJobStateFailedNotBlocking**|**6**||
|**pjCacheJobStateInvalid**|**-1**||
|**pjCacheJobStateLastState**|**13**||
|**pjCacheJobStateOnHold**|**10**||
|**pjCacheJobStateProcessing**|**3**||
|**pjCacheJobStateReadyForLaunch**|**12**||
|**pjCacheJobStateReadyForProcessing**|**1**||
|**pjCacheJobStateSendIncomplete**|**2**||
|**pjCacheJobStateSkipped**|**7**||
|**pjCacheJobStateSleeping**|**11**||
|**pjCacheJobStateSuccess**|**4**||
|**pjCacheJobStateUnknown**|**0**||

## Remarks

The  **[Application.GetCacheStatusForProject](application-getcachestatusforproject-property-project.md)** property returns a **PjCacheJobState** constant. The **PjCacheJobState** constants from 0 to 13 correspond to the[JobState enumeration](http://msdn.microsoft.com/en-us/library/websvcqueuesystem.jobstate_di_pj14mref%28office.15%29.aspx) of the **QueueSystem** service in the Project Server Interface (PSI).


## See also


#### Other resources


[GetCacheStatusForProject Property](application-getcachestatusforproject-property-project.md)
[PjJobType Enumeration](pjjobtype-enumeration-project.md)
[QueueConstants.JobState enumeration](http://msdn.microsoft.com/en-us/library/microsoft.office.project.server.library.queueconstants.jobstate_di_pj14mref%28office.15%29.aspx)
