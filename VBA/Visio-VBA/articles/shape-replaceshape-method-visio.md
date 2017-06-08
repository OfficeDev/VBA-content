---
title: Shape.ReplaceShape Method (Visio)
ms.prod: visio
ms.assetid: b330a63d-4e3f-0c4d-c38c-6ee806670225
ms.date: 06/08/2017
---


# Shape.ReplaceShape Method (Visio)

Replaces the specified shape with an instance of the master passed as the first parameter, and returns the new shape.


## Syntax

 _expression_ . **ReplaceShape**_(MasterOrMasterShortcutToDrop,_ _ReplaceFlags)_

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _MasterOrMasterShortcutToDrop_|Required|UNKNOWN|Specifies the replacement shape to drop. Must be either a [Master](master-object-visio.md) or[MasterShortcut](mastershortcut-object-visio.md) object.|
| _ReplaceFlags_|Optional|INT32|Specifies the properties of the original shape to retain in the new shape. Possible values include any of the [VisReplaceFlags](visreplaceflags-enumeration-visio.md) constants, and certain combinations of those constants. See Remarks for more information.|

### Return value

 **SHAPE**


## Remarks

Allowable values to pass for the  _ReplaceFlags_ parameter include either **visReplaceShapeDefault** or any combination of one or more of the remaining four flags.


## See also


#### Concepts


[Shape Object](shape-object-visio.md)

