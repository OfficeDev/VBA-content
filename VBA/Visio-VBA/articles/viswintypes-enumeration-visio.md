---
title: VisWinTypes Enumeration (Visio)
keywords: vis_sdr.chm70005
f1_keywords:
- vis_sdr.chm70005
ms.prod: visio
ms.assetid: 9d5ecb3f-baf8-8d9b-608a-8b9661b04ec9
ms.date: 06/08/2017
---


# VisWinTypes Enumeration (Visio)

Type and ID codes returned by  **Window.Type** , **Window.SubType** , and **Window.ID** .



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|visAnchorBarAddon|10|Window created by an add-on that has tabs at the bottom when merged (floating, anchored, or docked window)|
|visAnchorBarBuiltIn|6|Visio built-in window that has tabs at the bottom when merged?presently, the  **Custom Properties, Size &; Position, Drawing Explorer, Master Explorer, Pan &; Zoom and Validation Issues** windows (floating, anchored, or docked windows).|
|visApplication|5|Microsoft Visio application window.|
|visDockedStencilAddon|11|An add-on window that has docked stencil behavior.|
|visDockedStencilBuiltIn|7|Stencil window docked in a drawing window.|
|visDrawing|1|Drawing window (MDI frame window).|
|visDrawingAddon|8|Drawing window created by an add-on (MDI frame window).|
|visIcon|4|Icon editing window (MDI frame window).|
|visInvalWinID|-1|Window has no ID.|
|visMasterGroupWin|96|A group editing window of a group in a master.|
|visMasterWin|64|A master drawing page window.|
|visPageGroupWin|160|A group editing window of a group on a page.|
|visPageWin|128|A drawing window showing a page.|
|visSheet|3|ShapeSheet window (MDI frame window).|
||2|Stencil window (MDI frame window).|
|visStencil|9|Add-on window that has stencil window behavior.|
|visStencilAddon|1658|When  **Window.Type** is **visAnchorBarBuiltIn** , **Custom Properties window** .|
|visWinIDCustProp|1721|When  **Window.Type** is **visAnchorBarBuiltIn** , **Drawing Explorer** window.|
|visWinIDDrawingExplorer|2044|When  **Window.Type** is **visAnchorBarBuiltIn** , **External Data window** .|
|visWinIDExternalData|1781|When  **Window.Type** is **visAnchorBarBuiltIn** , ShapeSheet **Formula Tracing** window.|
|visWinIDFormulaTracing|1916|When  **Window.Type** is **visAnchorBarBuiltIn** , **Master Explorer** window in master editing window.|
|visWinIDMasterExplorer|1653|When  **Window.Type** is **visAnchorBarBuiltIn** , **Pan &; Zoom** window.|
|visWinIDPanZoom|1669|When  **Window.Type** is **visAnchorBarBuiltIn** , **Shapes** window.|
|visWinIDSizePos|1670|When  **Window.Type** is **visAnchorBarBuiltIn** , **Size &; Position** window.|
|visWinIDStencilExplorer|1796|When  **Window.Type** is **visAnchorBarBuiltIn** , **Drawing Explorer** window in MDI stencil window.|
|visWinIDValidationIssues|2263|When  **Window.Type** is **visAnchorBarBuiltIn** , **Validation Issues** window.|
|visWinOther|0|Unknown window type.|

