---
title: PjConstraint Enumeration (Project)
ms.prod: project-server
api_name:
- Project.PjConstraint
ms.assetid: 1ba4f126-18b8-0c74-a26d-d896ca5f87dd
ms.date: 06/08/2017
---


# PjConstraint Enumeration (Project)

Contains constants that specify the type of constraint.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
|**pjALAP**|1|Task occurs as late as possible in the schedule without delaying subsequent tasks.|
|**pjASAP**|0|Task occurs as soon as possible in the schedule. This is the default constraint type for tasks.|
|**pjFNET**|6|Task finishes on or after the constraint date.|
|**pjFNLT**|7|Task finishes on or before the constraint date.|
|**pjMFO**|3|Task finishes on the constraint date.|
|**pjMSO**|2|Task starts on the constraint date.|
|**pjSNET**|4|Task starts on or after the constraint date.|
|**pjSNLT**|5|Task starts on or before the constraint date.|

