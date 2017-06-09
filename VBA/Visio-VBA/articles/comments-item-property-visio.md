---
title: Comments.Item Property (Visio)
ms.prod: visio
ms.assetid: fed2a079-de87-d5ce-1d74-0bfa5a328441
ms.date: 06/08/2017
---


# Comments.Item Property (Visio)

Returns an object from a collection. The  **Item** property is the default property for all collections. Read-only[Comment](comment-object-visio.md).


## Syntax

 _expression_ . **Item**

 _expression_ A variable that represents a **Comments** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|Contains the index of the object to retrieve.|

### Return Value

 **Comment**


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:


```
objRet = object(index )
```


## Property value

 **IVCOMMENT**


## See also


#### Other resources


[Comments Collection](comments-object-visio.md)

