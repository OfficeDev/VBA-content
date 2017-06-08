---
title: Remove Method (FileSystemObject object)
keywords: vblr6.chm2181952
f1_keywords:
- vblr6.chm2181952
ms.prod: office
ms.assetid: dc895fae-17aa-4c51-4a35-8c3d3fd0e6fc
ms.date: 06/08/2017
---


# Remove Method (FileSystemObject object)



 **Description**
Removes a key, item pair from a  **Dictionary** object.
 **Syntax**
 _object_. **Remove(**_key_**)**
The  **Remove** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **Dictionary** object.|
| _key_|Required.  _Key_ associated with the key, item pair you want to remove from the **Dictionary** object.|
 **Remarks**
An error occurs if the specified key, item pair does not exist.
The following code illustrates use of the  **Remove** method:



```vb
Dim a, d, i             'Create some variables
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", "Athens"     'Add some keys and items
d.Add "b", "Belgrade"
d.Add "c", "Cairo"
...
a = d. Remove()          'Remove second pair

```


