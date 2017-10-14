---
title: Event.ID Property (Visio)
keywords: vis_sdr.chm12613675
f1_keywords:
- vis_sdr.chm12613675
ms.prod: visio
api_name:
- Visio.Event.ID
ms.assetid: d1c5ae17-eb31-48c7-f63a-02121d44f6f5
ms.date: 06/08/2017
---


# Event.ID Property (Visio)

Gets the ID of an object. Read-only.


## Syntax

 _expression_ . **ID**

 _expression_ A variable that represents an **Event** object.


### Return Value

Long


## Remarks

The ID of a shape is unique only within the scope of the page or master. The ID of a page, master, or style is unique within the scope of the document.

If a shape, page, master, or style is deleted, future objects in the same scope may be assigned the same ID. Therefore, persisting shape or style IDs in separate data stores is generally not as sound as persisting unique IDs using the  **UniqueID** property.

For  **Shape** objects, you can use the **ID** property with methods such as **GetResults** and **SetResults** to get or set many cell values at once, possibly cells in many different shapes. To do this, you must pass shape IDs to the methods. If you create shapes by using the **DropMany** method, the method returns the IDs of the shapes it creates to your program.

For  **Font** objects, the **ID** property corresponds to the number stored in the Font cell of the row in a shape's Character Properties section. For example, to apply the font named "Arial" to a shape's text, create a **Font** object that represents "Arial," get the ID of that font, and then set the **CharProps** property of the **Shape** object to that ID.

The ID associated with a particular font varies from system to system or as fonts are installed and removed on a given system.

For  **Window** objects, the **ID** property can be used with the **ItemFromID** property of a **Windows** collection to retrieve a **Window** object from the collection without iterating through the collection. A **Window** object whose **Type** property is set to **visAnchorBarBuiltIn** returns an ID of **visWinIDCustProp** , **visWinIDDrawingExplorer** , **visWinIDFormulaTracing** , **visWinIDMasterExplorer** , **visWinIDPanZoom** , **visWinIDSizePos** , or **visWinIDStencilExplorer** . A **Window** object whose **Type** property is set to **visAnchorBarAddon** returns an ID that is unique within its **Windows** collection for the lifetime of that collection. If a **Window** object has an ID of **visInvalWinID** , you cannot use the **ItemFromID** property to retrieve the **Window** object from its collection.

For  **Event** objects, the **ID** property uniquely identifies an **Event** object in its **EventList** collection. As long as a reference is held on an **EventList** collection or on the source object of an **EventList** collection, you can cache the **ID** property of any **Event** object in the list. Even if other events are added to or removed from the list, the cached ID can be used later to identify the original event. If an event is persistent, its ID can be cached indefinitely. While the event with that ID might be removed, no new **Event** object in the same **EventList** collection is given the same ID.


