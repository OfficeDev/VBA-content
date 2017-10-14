---
title: GraphicItem Object (Visio)
keywords: vis_sdr.chm61035
f1_keywords:
- vis_sdr.chm61035
ms.prod: visio
api_name:
- Visio.GraphicItem
ms.assetid: 80b4b4da-9ed2-dcbc-8f96-70f1b07c2b20
ms.date: 06/08/2017
---


# GraphicItem Object (Visio)

Represents a single component part of a data graphic master (a  **[Master](master-object-visio.md)** object of type **visTypeDataGraphic** ) that is responsible for a specific graphical adornment of the master.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Remarks

The default property of a  **GraphicItem** object is **[ID](graphicitem-id-property-visio.md)** .

 **GraphicItem** objects can be of four types:




- Color by value
    
- Data bar
    
- Icon set
    
- Text callout
    


You cannot create or significantly modify  **GraphicItem** objects programmatically—you must perform these tasks in the Visio user interface (UI). For more information about creating and modifying data graphics in the UI, search for "data graphic" in Visio end-user Help.

However, the properties and methods listed below do permit some programmatic modification of  **GraphicItem** objects. In particular, you can modify the position of a graphic item relative to the shape or selection it's associated with; the Z-order (the order in which Visio draws graphic items) of a **GraphicItem** object in the **[GraphicItems](graphicitems-object-visio.md)** collection; and the expression (value) used to evaluate a graphic item against the rule that determines how it is displayed.

Use the  **[DataGraphic](graphicitem-datagraphic-property-visio.md)** property to get the **Master** object of type **visTypeDataGraphic** that contains the **GraphicItem** object.

Use the  **[GetExpression](graphicitem-getexpression-method-visio.md)** method to get the current expression against which the graphic item's rule is evaluated.

Use the  **[SetExpression](graphicitem-setexpression-method-visio.md)** method to set the current expression against which the graphic item's rule is evaluated.

Use the  **[Delete](graphicitem-delete-method-visio.md)** method to delete a **GraphicItem** object from the **[GraphicItems](graphicitems-object-visio.md)** collection.

Use the  **[HorizontalPosition](graphicitem-horizontalposition-property-visio.md)** property to get or set the horizontal position of the graphic item relative to the shape or selection that it's associated with.

Use the  **[VerticalPosition](graphicitem-verticalposition-property-visio.md)** property to get or set thevertical position of the graphic item relative to the shape or selection that it's associated with.

Use the  **[UseDataGraphicPosition](graphicitem-usedatagraphicposition-property-visio.md)** property to get or set whether a **GraphicItem** object inherits the settings of the **[DataGraphicHorizontalPosition](master-datagraphichorizontalposition-property-visio.md)** and **[DataGraphicVerticalPosition](master-datagraphicverticalposition-property-visio.md)** properties of the data graphic master it belongs to (when set to **True)** , or whether the **GraphicItem** object's own **HorizontalPosition** and **Vertical Position** settings are applied (when set to **False** ).


