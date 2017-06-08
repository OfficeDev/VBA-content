---
title: InvisibleApp.QueryCancelReplaceShapes Event (Visio)
ms.prod: visio
ms.assetid: 5e5d9b76-dfd4-1d02-d205-9e64350449d5
ms.date: 06/08/2017
---


# InvisibleApp.QueryCancelReplaceShapes Event (Visio)

Occurs immediately after a shape-replacement operation is requested. If any event handler returns  **True** , the operation is canceled.


## Syntax

 _expression_ . **QueryCancelReplaceShapes**( _replaceShapes_)

 _expression_ A variable that represents a **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _replaceShapes_|Required|REPLACESHAPESEVENT|An object whose properties return information about the shape-replacement operation.|
|||||
| _lpboolRet_|Required|BOOL||

## See also


#### Concepts


[InvisibleApp Object](invisibleapp-object-visio.md)

