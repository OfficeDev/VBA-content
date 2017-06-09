---
title: Shape.MasterShape Property (Visio)
keywords: vis_sdr.chm11250710
f1_keywords:
- vis_sdr.chm11250710
ms.prod: visio
api_name:
- Visio.Shape.MasterShape
ms.assetid: bf710d8b-11f6-145d-a306-658dc23dedbf
ms.date: 06/08/2017
---


# Shape.MasterShape Property (Visio)

If this shape is part of a master instance, returns the shape in the master that this shape inherits from. Read-only.


## Syntax

 _expression_ . **MasterShape**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Shape


## Remarks

Each shape in an instance of a master (the group and all its subshapes) points to its corresponding shape in the master. The  **MasterShape** property returns the **Shape** object in the master from which the parent **Shape** object inherits.

If the parent  **Shape** object is not part of a master instance, the **MasterShape** property returns **Nothing** .


