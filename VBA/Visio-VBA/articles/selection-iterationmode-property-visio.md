---
title: Selection.IterationMode Property (Visio)
keywords: vis_sdr.chm11113785
f1_keywords:
- vis_sdr.chm11113785
ms.prod: visio
api_name:
- Visio.Selection.IterationMode
ms.assetid: e4cd372c-a156-364d-f051-d9a8c618bd2c
ms.date: 06/08/2017
---


# Selection.IterationMode Property (Visio)

Determines whether a  **Selection** object reports subselected shapes and groups in which a shape is selected. Read/write.


## Syntax

 _expression_ . **IterationMode**

 _expression_ A variable that represents a **Selection** object.


### Return Value

Long


## Remarks

The items in a  **Selection** object are a subset of the descendants of the **Selection** object's containing shape.




- A top-level shape in a  **Selection** object is an immediate child of the selection's containing shape.
    
- A subselected shape in a  **Selection** object is not an immediate child of the selection's containing shape.
    
- A superselected shape in a  **Selection** object has at least one immediate child that is subselected.
    


If a shape is subselected, each of its ancestors?except the containing shape itself?is superselected.

The value of the  **IterationMode** property is a combination of the following values.



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visSelModeSkipSuper**|&;H0100 |Selection does not report superselected shapes. |
| **visSelModeOnlySuper**|&;H0200 |Selection only reports superselected shapes. |
| **visSelModeSkipSub**|&;H0400 |Selection does not report subselected shapes. |
| **visSelModeOnlySub**|&;H0800 |Selection only reports subselected shapes. |
When a  **Selection** object is created, its initial iteration mode is **visSelModeSkipSub** + **visSelModeSkipSuper** . It reports neither subselected nor superselected shapes and behaves identically to **Selection** objects in versions of Microsoft Visio prior to Visio 2000.

You can determine whether an individual item in a  **Selection** object is a subselected or superselected item by using the **ItemStatus** property.


