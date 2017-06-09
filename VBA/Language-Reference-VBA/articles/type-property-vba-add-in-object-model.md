---
title: Type Property (VBA Add-In Object Model)
keywords: vbob6.chm100096
f1_keywords:
- vbob6.chm100096
ms.prod: office
ms.assetid: 358eca6e-05fd-2c01-f004-b80bd909338e
ms.date: 06/08/2017
---


# Type Property (VBA Add-In Object Model)



Returns a numeric or string value containing the type of object. Read-only.
 **Return Values**
The  **Type** property settings for the **Window** object are described in the following table:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbext_wt_CodeWindow**|0|**Code** window|
|**vbext_wt_Designer**|1|[Designer](vbe-glossary.md)|
|**vbext_wt_Browser**|2|**Object Browser**|
|**vbext_wt_Immediate**|5|**Immediate** window|
|**vbext_wt_ProjectWindow**|6|[Project window](vbe-glossary.md)|
|**vbext_wt_PropertyWindow**|7|[Properties window](vbe-glossary.md)|
|**vbext_wt_Find**|8|**Find** dialog box|
|**vbext_wt_FindReplace**|9|**Search and Replace** dialog box|
|**vbext_wt_LinkedWindowFrame**|11|[Linked window frame](vbe-glossary.md)|
|**vbext_wt_MainWindow**|12|Main window|
|**vbext_wt_Watch**|3|**Watch** window|
|**vbext_wt_Locals**|4|**Locals** window|
|**vbext_wt_Toolbox**|10|**Toolbox**|
|**vbext_wt_ToolWindow**|15|**Tool** window|


 **Important**  Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements will generate run-time errors when run on the Macintosh.


The  **Type** property settings for the **VBComponent** object are described in the following table:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbext_ct_StdModule**|1|[Standard module](vbe-glossary.md)|
|**vbext_ct_ClassModule**|2|[Class module](vbe-glossary.md)|
|**vbext_ct_MSForm**|3|Microsoft Form|
|**vbext_ct_ActiveXDesigner**|11|ActiveX Designer|
|**vbext_ct_Document**|100|Document Module|
The  **Type** property settings for the **Reference** object are described in the following table:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbext_rk_TypeLib**|0|[Type library](vbe-glossary.md)|
|**vbext_rk_Project**|1|[Project](vbe-glossary.md)|
The  **Type** property settings for the **VBProject** object are described in the following table:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbext_pt_HostProject**|100|Host Project|
|**vbext_pt_StandAlone**|101|Standalone Project|

