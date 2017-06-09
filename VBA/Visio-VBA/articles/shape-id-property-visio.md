---
title: Shape.ID Property (Visio)
keywords: vis_sdr.chm11213675
f1_keywords:
- vis_sdr.chm11213675
ms.prod: visio
api_name:
- Visio.Shape.ID
ms.assetid: 948982c0-a872-802f-a2d3-69c6539ca3f2
ms.date: 06/08/2017
---


# Shape.ID Property (Visio)

Gets the ID of an object. Read-only.


## Syntax

 _expression_ . **ID**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Long


## Remarks

The ID of a shape is unique only within the scope of the page or master. 

If a shape, page, master, or style is deleted, future objects in the same scope may be assigned the same ID. Therefore, persisting shape or style IDs in separate data stores is generally not as sound as persisting unique IDs using the  **UniqueID** property.

For  **Shape** objects, you can use the **ID** property with methods such as **GetResults** and **SetResults** to get or set many cell values at once, possibly cells in many different shapes. To do this, you must pass shape IDs to the methods. If you create shapes by using the **DropMany** method, the method returns the IDs of the shapes it creates to your program.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVShape.ID**
    

