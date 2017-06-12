---
title: Master.IndexInStencil Property (Visio)
keywords: vis_sdr.chm10713700
f1_keywords:
- vis_sdr.chm10713700
ms.prod: visio
api_name:
- Visio.Master.IndexInStencil
ms.assetid: 3c2c12c4-0233-4aa3-c3d7-a3613bb391ad
ms.date: 06/08/2017
---


# Master.IndexInStencil Property (Visio)

Gets or sets the index of a master or master shortcut object within its stencil. Read/write.


## Syntax

 _expression_ . **IndexInStencil**

 _expression_ A variable that represents a **Master** object.


### Return Value

Integer


## Remarks

Beginning with Visio 2000, the document stencil window shows all  **Master** and **MasterShortcut** objects in a Visio document. The Visio object model exposes the **Master** and **MasterShortcut** objects in a **Document** object as two distinct collections. The index returned by a **Master** object is its index with respect to other **Master** objects in its **Document** object and is unrelated to the presence or absence of **MasterShortcut** objects in the document. The index returned by a **MasterShortcut** object is its index with respect to other **MasterShortcut** objects in its **Document** object and is unrelated to the presence or absence of **Master** objects in the document.

Use the  **IndexInStencil** property to maintain the relative order of **Master** and **MasterShortcut** objects when considered as a single collection.




 **Note**  Beginning with Microsoft Office Visio 2003, only user-created stencils are editable. By default, Visio stencils are not editable. 


