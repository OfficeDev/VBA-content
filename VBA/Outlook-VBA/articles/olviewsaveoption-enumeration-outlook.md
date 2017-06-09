---
title: OlViewSaveOption Enumeration (Outlook)
keywords: vbaol11.chm3095
f1_keywords:
- vbaol11.chm3095
ms.prod: outlook
api_name:
- Outlook.OlViewSaveOption
ms.assetid: c08bab4d-ecdd-a2ac-1cdc-fa910f9585e0
ms.date: 06/08/2017
---


# OlViewSaveOption Enumeration (Outlook)

Specifies the folders in which the view is available and the read permissions attached to the view.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olViewSaveOptionAllFoldersOfType**|2|Indicates that the view is available in all folders of the same type.|
| **olViewSaveOptionThisFolderEveryone**|0|Indicates that the view is only available in the current folder and is available to all users.|
| **olViewSaveOptionThisFolderOnlyMe**|1|Indicates that the view is only available in the current folder and is only available to the current Outlook user.|

## Remarks

Used by the  **Copy** method and **SaveOption** property of **View** objects.


