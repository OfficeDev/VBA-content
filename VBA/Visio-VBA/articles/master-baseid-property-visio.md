---
title: Master.BaseID Property (Visio)
keywords: vis_sdr.chm10713135
f1_keywords:
- vis_sdr.chm10713135
ms.prod: visio
api_name:
- Visio.Master.BaseID
ms.assetid: 85ca3c0d-5015-b303-7102-144768acb6a8
ms.date: 06/08/2017
---


# Master.BaseID Property (Visio)

Returns a base ID for a master. Read-only.


## Syntax

 _expression_ . **BaseID**

 _expression_ A variable that represents a **Master** object.


### Return Value

String


## Remarks

A base ID is assigned to a master when it is created. When a master is copied, the copies all have the same base ID as the original master.

A  **Master** object also has a **UniqueID** property. When a master is copied, the copy is assigned the same unique ID as the original master, and its base also ID remains the same as that of the original master. If the copy of the master gets changed, its unique ID changes its the base ID remains the same.

In addition, if you copy into a stencil a master that has the same unique ID as a master already in the stencil, Visio assigns a new unique ID to the copy.

The only way to change a master's base ID is to use the  **NewBaseID** property.

If you know the base ID of a master, you can use the following code to retrieve the master from the  **Masters** collection of the active document:




```vb
'Retrieve the master whose BaseID value is 
'{0478DA94-1315-9876-8E4C-006523ABC9B2} 
Dim vsoMaster As Visio.Master 
Set vsoMaster = Visio.ActiveDocument.Masters("B{0478DA94-1315-9876-8E4C-006523ABC9B2}") 

```

If you know the base ID or the unique ID of a master, but are not sure which kind of ID it is, you can use the following code to retrieve the master from the  **Masters** collection of the active document:




```vb
'Retrieve the master whose UniqueID or BaseID value is 
'{0478DA94-1315-9876-8E4C-006523ABC9B2} 
Dim vsoMaster As Visio.Master 
Set vsoMaster = Visio.ActiveDocument.Masters("A{0478DA94-1315-9876-8E4C-006523ABC9B2}")
```


