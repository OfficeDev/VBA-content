---
title: Selection.ReplaceShape Method (Visio)
ms.prod: visio
ms.assetid: dc278901-77ce-e1fe-c44f-f464bbb1c360
ms.date: 06/08/2017
---


# Selection.ReplaceShape Method (Visio)

Replaces the specified selection with one or more instances of the master passed as the first parameter, and returns an array containing the new shape or shapes.


## Syntax

 _expression_ . **ReplaceShape**_(MasterOrMasterShortcutToDrop,_ _ReplaceFlags)_

 _expression_ A variable that represents a **Selection** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _MasterOrMasterShortcutToDrop_|Required|UNKNOWN|Specifies the replacement shape or shapes to drop. Must be either a [Master](master-object-visio.md) or[MasterShortcut](mastershortcut-object-visio.md) object.|
| _ReplaceFlags_|Optional|INT32|Specifies the properties of the original shape or shapes to retain in the new shape or shapes. Possible values include any of the [VisReplaceFlags](visreplaceflags-enumeration-visio.md) constants, and certain combinations of those constants. See Remarks for more information.|

### Return value

 **SAFE-ARRAY**


### Remarks

Allowable values to pass for the  _ReplaceFlags_ parameter include either **visReplaceShapeDefault** or any combination of one or more of the remaining four flags.


## See also


#### Concepts


[Selection Object](selection-object-visio.md)

