---
title: Viewer Object (Visio Viewer)
ms.prod: visio
ms.assetid: 4d25251a-5c4d-42d4-a73e-7e1e987ff593
ms.date: 06/08/2017
---


# Viewer Object (Visio Viewer)

The  **Viewer** object is a programmable ActiveX control that enables you to display Visio drawings (with limited functionality) on web pages and in Windows Forms, so that users who do not have Visio installed on their computers can view and interact with them.


## Remarks

With Visio Viewer, users can open, view, or print Visio drawings, even if they do not have Microsoft Visio 2013 installed. They cannot, however, edit or save drawings, or create a new Visio drawing. For that, they need to install Visio.

The  **Viewer** object is the entry point to the **Viewer** object model, and represents an instance of the Viewer control. The properties, events, and methods available in the **Viewer** object model let you load and unload Visio drawings in Visio Viewer, temporarily change properties and settings of the drawing, react to user input, and customize the Visio Viewer environment. In many cases, these members correspond to the options available to users in the Visio Viewer user interface (UI).

The following is a partial listing of the members of the  **Viewer** object and their functions and provides a sampling of the programming options available to developers. See the table of contents of this reference for the complete list of members. See [About Programming Visio Viewer](about-programming-visio-viewer.md) for code samples that show how to get an instance of the **Viewer** object in the available development environments.

Use the  **[Load](viewer-load-method-visio-viewer.md)** method to load a Visio drawing into Visio Viewer, and use the **[Unload](viewer-unload-method-visio-viewer.md)** method to unload the drawing. You can also use the **[SRC](viewer-src-property-visio-viewer.md)** property to get and set the file name and path for the current drawing.

Use the  **[DisplayAbout](viewer-displayabout-method-visio-viewer.md)**,  **[DisplayContextMenu](viewer-displaycontextmenu-method-visio-viewer.md)**,  **[DisplayHelp](viewer-displayhelp-method-visio-viewer.md)**, and  **[DisplayPropertyDialog](viewer-displaypropertydialog-method-visio-viewer.md)** methods to display the dialog boxes and shortcut menus available in the Visio Viewer UI.

Use the  **[SelectShape](viewer-selectshape-method-visio-viewer.md)** method to select a particular shape in the drawing and the **[ShapeName](viewer-shapename-property-visio-viewer.md)** and **[ShapeCount](viewer-shapecount-property-visio-viewer.md)** properties to get information about shapes in the drawing.

Use properties such as  **[BackColor](viewer-backcolor-property-visio-viewer.md)**,  **[GridVisible](viewer-gridvisible-property-visio-viewer.md)**,  **[LayerColor](viewer-layercolor-property-visio-viewer.md)**,  **[PageColor](viewer-pagecolor-property-visio-viewer.md)**,  **[ScrollbarsVisible](viewer-scrollbarsvisible-property-visio-viewer.md)**, and  **[ToolbarVisible](viewer-toolbarvisible-property-visio-viewer.md)** to customize the appearance of the Visio Viewer UI.

Use the  **[CustomPropertyCount](viewer-custompropertycount-property-visio-viewer.md)**,  **[CustomPropertyName](viewer-custompropertyname-property-visio-viewer.md)**, and  **[CustomPropertyValue](viewer-custompropertyvalue-property-visio-viewer.md)** properties to determine shape data (custom properties).

Use events such as  **[OnLayerChanged](viewer-onlayerchanged-event-visio-viewer.md)** and **[OnSelectionChanged](viewer-onselectionchanged-event-visio-viewer.md)** to respond to user input.


