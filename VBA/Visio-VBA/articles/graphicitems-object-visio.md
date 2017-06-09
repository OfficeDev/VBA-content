---
title: GraphicItems Object (Visio)
keywords: vis_sdr.chm61030
f1_keywords:
- vis_sdr.chm61030
ms.prod: visio
api_name:
- Visio.GraphicItems
ms.assetid: 89d0bbeb-ee45-50cc-490e-0af49d036ad1
ms.date: 06/08/2017
---


# GraphicItems Object (Visio)

The collection of  **[GraphicItem](graphicitem-object-visio.md)** objects associated with a **Master** object of type **visTypeDataGraphic** .


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Remarks

The default property of the  **GraphicItems** collection is **[Item](graphicitems-item-property-visio.md)** .

The  **[AddCopy](graphicitems-addcopy-method-visio.md)** method adds a copy of an existing **GraphicItem** object to the **GraphicItems** collection. The **GraphicItem** object to be added must already exist in another master of type **visTypeDataGraphic** .

Use the  **[DataGraphic](graphicitem-datagraphic-property-visio.md)** property to return the **Master** object of type **visTypeDataGraphic** that the **GraphicItems** collection is associated with.


 **Note**  You must create masters of type  **visTypeDataGraphic** by using the Visio user interface—you cannot create them programmatically. For more information about creating these masters, search for "data graphics" in Visio end-user Help.


