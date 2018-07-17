---
title: Broadcast.End Method (PowerPoint)
keywords: vbapp10.chm732004
f1_keywords:
- vbapp10.chm732004
ms.prod: powerpoint
api_name:
- PowerPoint.Broadcast.End
ms.assetid: b4ccda97-aa08-77ff-3a2f-8600721a55b0
ms.date: 06/08/2017
---


# Broadcast.End Method (PowerPoint)

Elevates to the system to delete the document from the Broadcast Documents library. 


## Syntax

 _expression_. **End**

 _expression_ A variable that represents a **Broadcast** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Required| _FileID_|**String**|The file to be removed from the Broadcast Documents library.|

### Return Value

None


## Remarks

This method validates that the user who made the request is the creator of the document. Elevation is necessary because presenters do not have access to directly delete documents from the Broadcast Documents library.


## See also


#### Concepts


[Broadcast Object](broadcast-object-powerpoint.md)

