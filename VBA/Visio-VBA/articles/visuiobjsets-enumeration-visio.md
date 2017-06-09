---
title: VisUIObjSets Enumeration (Visio)
keywords: vis_sdr.chm70170
f1_keywords:
- vis_sdr.chm70170
ms.prod: visio
ms.assetid: b5638672-73ba-aeb2-6660-bb44107f7ac8
ms.date: 06/08/2017
---


# VisUIObjSets Enumeration (Visio)

UI object-set identifiers, used with the  **SetID** and **ItemAtID** properties and the **AddAtID** method of the **AccelTable** , **MenuSet** , and **ToolbarSet** collections. Valid identifiers for particular collections are indicated.


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.



|**Constant**|**Value**|**Description**|**AccelTable**|**MenuSet**|**ToolbarSet**|
|:-----|:-----|:-----|:-----|:-----|:-----|
| **visUIObjSetActiveXDoc**|18|Visio is running as an ActiveX document.|X|||
| **visUIObjSetCntx_AddComments**|1000|Built-in commands available in the  **Customize** dialog box||X||
| **visUIObjSetCntx_AnchorBar_Base**|61|Shortcut menu: anchored windows.||X||
| **visUIObjSetCntx_AnchorBar_CustProp**|62|Shortcut menu:  **Shape Data** anchored window.||X||
| **visUIObjSetCntx_AnchorBar_Shapes**|69|Shortcut menu:  **Shapes** window.||||
| **visUIObjSetCntx_AnchorBar_SizePos**|63|Shortcut menu:  **Size &; Position** anchored window.||X||
| **visUIObjSetCntx_BuiltInMenus**|1010|Built-in menus available in the  **Customize** dialog box.||X||
| **visUIObjSetCntx_CommentMarker**|68|Shortcut menu:  **Comment Marker**.||X||
| **visUIObjSetCntx_ConnectPtType**|44|Shortcut menu:  **Connection Point**.||X||
| **visUIObjSetCntx_DEDocument**|49|Shortcut menu:  **Drawing Explorer**, root item.||X||
| **visUIObjSetCntx_DELayers**|54|Shortcut menu:  **Drawing Explorer**,  **Layers** folder.||X||
| **visUIObjSetCntx_DELayer**|55|Shortcut menu:  **Drawing Explorer**,  **Layer** item.||X||
| **visUIObjSetCntx_DEMasters**|58|Shortcut menu:  **Drawing Explorer**,  **Masters** folder.||X||
| **visUIObjSetCntx_DEMaster**|59|Shortcut menu:  **Drawing Explorer**, item.||X||
| **visUIObjSetCntx_DEPages**|50|Shortcut menu:  **Drawing Explorer**,  **Pages** folder.||X||
| **visUIObjSetCntx_DEPage**|51|Shortcut menu:  **Drawing Explorer**, item.||X||
| **visUIObjSetCntx_DEPatterns**|60|Shortcut menu:  **Drawing Explorer**,  **Patterns** folder.||X||
| **visUIObjSetCntx_DEShapes**|52|Shortcut menu:  **Drawing Explorer**,  **Shapes** folder.||X||
| **visUIObjSetCntx_DEShape**|53|Shortcut menu:  **Drawing Explorer**, item.||X||
| **visUIObjSetCntx_DEStyles**|56|Shortcut menu:  **Drawing Explorer**,  **Styles** folder.||X||
| **visUIObjSetCntx_DEStyle**|57|Shortcut menu:  **Drawing Explorer**,  **Style** item.||X||
| **visUIObjSetCntx_DrawObjSel**|9|Shortcut menu: Visio drawing shape.||X||
| **visUIObjSetCntx_DrawOleObjSel**|10|Shortcut menu: foreign drawing shape.||X||
| **visUIObjSetCntx_Fullscreen**|17|Shortcut menu:  **Full Screen** mode.||X||
| **visUIObjSetCntx_Issues**|76|Issues window.||||
| **visUIObjSetCntx_Master**|14|Shortcut menu:  **Master**.||X||
| **visUIObjSetCntx_MEDocument**|66|Shortcut menu:  **Master Explorer**, root item||X||
| **visUIObjSetCntx_MEMasters**|67|Shortcut menu:  **Master Explorer**,  **Masters** folder||X||
| **visUIObjSetCntx_Page**|75|Page in drawing window, no shape selected.||||
| **visUIObjSetCntx_PageTabNavigation**|74|Page tab navigation.||||
| **visUIObjSetCntx_PageTabs**|47|Shortcut menu:  **Page** tab||X||
| **visUIObjSetCntx_ShapeSheet**|15|Shortcut menu: Shape window.||X||
| **visUIObjSetCntx_Stencil**|21|Shortcut menu: Stencil window.||X||
| **visUIObjSetCntx_TextEdit**|13|Shortcut menu: editing shape text.||X||
| **visUIObjSetDrawing**|2|Drawing window is active or no documents are open.|X|X|X|
| **visUIObjSetHostingInPlace**|22|An object is active in Visio.|X|||
| **visUIObjSetIcon**|5|Icon editing window is active.|X|X|X|
| **visUIObjSetInPlace**|6|Use for accelerators only when Visio is running in place.|X|||
| **visUIObjSetPal_AlignShapes**|30|Toolbar palette:  **Align Shapes**|||X|
| **visUIObjSetPal_ConnectorTool**|40|Toolbar palette:  **Connector**,  **Connection Point**, and  **Stamp** drawing tools|||X|
| **visUIObjSetPal_CornerRounding**|34|Toolbar palette:  **Corner Rounding**|||X|
| **visUIObjSetPal_DistributeShapes**|31|Toolbar palette:  **Distribute Shapes**|||X|
| **visUIObjSetPal_FillColors**|27|Toolbar palette:  **Fill Color**|||X|
| **visUIObjSetPal_FillPatterns**|28|Toolbar palette:  **Fill Pattern**|||X|
| **visUIObjSetPal_LineColors**|24|Toolbar palette:  **Line Color**|||X|
| **visUIObjSetPal_LineEnds**|33|Toolbar palette:  **Line Ends**|||X|
| **visUIObjSetPal_LinePatterns**|26|Toolbar palette:  **Line Pattern**|||X|
| **visUIObjSetPal_LineTool**|42|Toolbar palette:  **Line**,  **Arc**,  **Pencil**, and  **Freeform** drawing tool|||X|
| **visUIObjSetPal_LineWeights**|25|Toolbar palette:  **Line Weight**|||X|
| **visUIObjSetPal_Rectangle_Tool**|37|Toolbar palette:  **Rectangle** and **Ellipse** drawing tools|||X|
| **visUIObjSetPal_RotationTool**|43|Toolbar palette:  **Rotation** and **Crop** tools|||X|
| **visUIObjSetPal_Shadow**|32|Toolbar palette:  **Shadow Color**|||X|
| **visUIObjSetPal_TextColors**|29|Toolbar palette:  **Text Color**|||X|
| **visUIObjSetPal_TextTool**|41|Toolbar palette:  **Text** and **Text Block** drawing tools|||X|
| **visUIObjSetPopup_LineJumpCode**|38|Toolbar palette: Add line jumps to|||X|
| **visUIObjSetPopup_LineJumpStyle**|39|Toolbar palette:  **Line Jump Style**|||X|
| **visUIObjSetPrintPreview**|7|Visio is in  **Print Preview** mode.|X|X|X|
| **visUIObjSetShapeSheet**|4|ShapeSheet window is active.|X|X|X|
| **visUIObjSetStencil**|3|Stencil window is active.|X|X|X|

