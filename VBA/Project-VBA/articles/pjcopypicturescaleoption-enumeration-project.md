---
title: PjCopyPictureScaleOption Enumeration (Project)
ms.prod: project-server
api_name:
- Project.PjCopyPictureScaleOption
ms.assetid: c9b995a6-67a4-93bb-6ed0-1a5f738db537
ms.date: 06/08/2017
---


# PjCopyPictureScaleOption Enumeration (Project)

Contains constants that specify how to treat a picture of the active view if it is larger than  **MaxImageWidth** by **MaxImageHeight**.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
|**pjCopyPictureKeepRange**|1|Maintains the selection, regardless of size. If the picture exceeds the amount of available memory, it is cropped to the maximum available size.|
|**pjCopyPictureScale**|2|Scales the picture to  **MaxImageWidth** by **MaxImageHeight MeasurementUnits** without maintaining the aspect ratio.|
|**pjCopyPictureScaleWRatio**|3| Scales the picture to **MaxImageWidth** by **MaxImageHeight MeasurementUnits** and maintains the aspect ratio.|
|**pjCopyPictureShowOptions**|0|Displays the  **Copy Picture Options** dialog box.|
|**pjCopyPictureTimescale**|4|Adjusts the timescale (zooms out) so that the picture fits  **MaxImageWidth** by **MaxImageHeight MeasurementUnits**.|
|**pjCopyPictureTruncate**|5|Truncates any portion of the picture that exceeds the boundaries of  **MaxImageWidth** by **MaxImageHeight MeasurementUnits**.|

