---
title: Item Property (FileSystemObject object)
keywords: vblr6.chm2181946
f1_keywords:
- vblr6.chm2181946
ms.prod: office
ms.assetid: 3f3080fa-29bf-6bf1-bd6e-6b804c01a589
ms.date: 06/08/2017
---


# Item Property (FileSystemObject object)



 **Description**
Sets or returns an  _item_ for a specified _key_ in a **Dictionary** object. For collections, returns an _item_ based on the specified _key_. Read/write.
 **Syntax**
 _object_. **Item(**_key_**)** [= _newitem_ ]
The  **Item** property has the following parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a collection or  **Dictionary** object.|
| _key_|Required.  _Key_ associated with the item being retrieved or added.|
| _newitem_|Optional. Used for  **Dictionary** object only; no application for collections. If provided, _newitem_ is the new value associated with the specified _key_.|
 **Remarks**
If  _key_ is not found when changing an _item_, a new _key_ is created with the specified _newitem_. If _key_ is not found when attempting to return an existing item, a new _key_ is created and its corresponding item is left empty.

