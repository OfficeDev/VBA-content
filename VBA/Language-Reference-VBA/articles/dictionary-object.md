---
title: Dictionary Object
keywords: vblr6.chm2181922
f1_keywords:
- vblr6.chm2181922
ms.prod: office
api_name:
- Office.Dictionary
ms.assetid: 718dbcd4-63bc-3a75-fa55-7d1e8c65e8b9
ms.date: 06/08/2017
---


# Dictionary Object



 **Description**
Object that stores data key, item pairs.
 **Syntax**
 **Scripting.Dictionary**
 **Remarks**
A  **Dictionary** object is the equivalent of a PERL associative array. Items, which can be any form of data, are stored in the array. Each item is associated with a unique key. The key is used to retrieve an individual item and is usually a integer or a string, but can be anything except an array.
The following code illustrates how to create a  **Dictionary** object:



```vb
Dim d                   'Create a variable
Set d = CreateObject(Scripting.Dictionary)
d.Add "a", "Athens"     'Add some keys and items
d.Add "b", "Belgrade"
d.Add "c", "Cairo"
...

```


