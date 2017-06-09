---
title: Master.UniqueID Property (Visio)
keywords: vis_sdr.chm10751165
f1_keywords:
- vis_sdr.chm10751165
ms.prod: visio
api_name:
- Visio.Master.UniqueID
ms.assetid: 99d0655c-da5c-9d0a-4936-2fa24821e097
ms.date: 06/08/2017
---


# Master.UniqueID Property (Visio)

Returns the unique ID of a master. Read-only.


## Syntax

 _expression_ . **UniqueID**

 _expression_ An expression that returns a **Master** object.


### Return Value

String


## Remarks

A  **Master** object always has a unique ID. If you copy a master, the new master has the same unique ID as the original master (as well as the same base ID). However, if you subsequently change the copy, Visio assigns it a new unique ID, but its base ID remains the same.

Note that if you copy into a stencil a master that has the same unique ID as a master already in the stencil, Visio assigns a new unique ID to the copy. 

For more information about the base ID, see the  **BaseID** property.

You can determine a  **Master** object's unique ID by using the following code:




```
strID = vsoMaster.UniqueID
```

The value it returns is a string in the following form:




```
{2287DC42-B167-11CE-88E9-0020AFDDD917}
```

To get a master if you know its unique ID, use  **Masters.Item** ( _UniqueIDString_) .

For example, you can use the following code to retrieve the master from the  **Masters** collection of the active document:




```vb
Dim vsoMaster As Visio.Master 
Set vsoMaster = Visio.ActiveDocument.Masters("{0478DA94-1315-9876-8E4C-006523ABC9B2}") 

```

Alternatively, you can use the following code, which adds the letter "U" before the string to identify it as a unique ID:




```vb
Dim vsoShape As Visio.Shape 
Set vsoMaster = Visio.ActiveDocument.Masters("U{0478DA94-1315-9876-8E4C-006523ABC9B2}") 

```


