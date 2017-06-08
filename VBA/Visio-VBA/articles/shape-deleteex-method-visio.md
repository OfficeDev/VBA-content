---
title: Shape.DeleteEx Method (Visio)
keywords: vis_sdr.chm11262730
f1_keywords:
- vis_sdr.chm11262730
ms.prod: visio
api_name:
- Visio.Shape.DeleteEx
ms.assetid: df4c164d-576a-acce-3322-7f166eb81e4f
ms.date: 06/08/2017
---


# Shape.DeleteEx Method (Visio)

Deletes the additional shapes that are associated with the shape, such as connectors and unselected container members, when the shape is deleted.


## Syntax

 _expression_ . **DeleteEx**( **_DelFlags_** )

 _expression_ A variable that represents a **[Shape](shape-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DelFlags_|Required| **Long**|The additional shapes to delete. See Remarks for possible values.|

### Return Value

 **Nothing**


## Remarks

 _DelFlags_ must be one or a bitwise combination of the following **[VisDeleteFlags](visdeleteflags-enumeration-visio.md)** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visDeleteNormal**|0|Match the deletion behavior that is in the user interface.|
| **visDeleteHealConnectors**|1|Delete connectors that are attached to deleted shapes.|
| **visDeleteNoHealConnectors**|2|Do not delete connectors that are attached to deleted shapes.|
| **visDeleteNoContainerMembers**|4|Do not delete unselected members of containers or lists.|
| **visDeleteNoAssociatedCallouts**|8|Do not delete unselected callouts that are associated with shapes.|
In a bitwise combination of  _DelFlags_ constants, you cannot combine **visDeleteHealConnectors** and **visDeleteNoHealConnectors** . If you attempt to do so, Microsoft Visio returns an Invalid Parameter error.


