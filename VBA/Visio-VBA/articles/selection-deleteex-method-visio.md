---
title: Selection.DeleteEx Method (Visio)
keywords: vis_sdr.chm11162730
f1_keywords:
- vis_sdr.chm11162730
ms.prod: visio
api_name:
- Visio.Selection.DeleteEx
ms.assetid: 8935a2de-2fab-0b2e-1595-a78d3dc2fd90
ms.date: 06/08/2017
---


# Selection.DeleteEx Method (Visio)

Deletes additional shapes associated with the selection, such as connectors and unselected container members, when the selection is deleted.


## Syntax

 _expression_ . **DeleteEx**( **_DelFlags_** )

 _expression_ A variable that represents a **[Selection](selection-object-visio.md)** object.


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
| **visDeleteNormal**|0|Match the deletion behavior in the user interface.|
| **visDeleteHealConnectors**|1|Delete connectors attached to deleted shapes.|
| **visDeleteNoHealConnectors**|2|Do not delete connectors attached to deleted shapes.|
| **visDeleteNoContainerMembers**|4|Do not delete unselected members of containers or lists.|
| **visDeleteNoAssociatedCallouts**|8|Do not delete unselected callouts associated with shapes.|
In a bitwise combination of  _DelFlags_ constants, you cannot combine **visDeleteHealConnectors** and **visDeleteNoHealConnectors** . If you attempt to do so, Microsoft Visio returns an Invalid Parameter error.


