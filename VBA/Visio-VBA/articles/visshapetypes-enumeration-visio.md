---
title: VisShapeTypes Enumeration (Visio)
keywords: vis_sdr.chm70010
f1_keywords:
- vis_sdr.chm70010
ms.prod: visio
ms.assetid: aa65da44-c6f4-bb14-e58b-43222fc066f5
ms.date: 06/08/2017
---


# VisShapeTypes Enumeration (Visio)

Shape type codes returned by the  **Shape.Type** and **Shape.ForeignType** properties.


## Remarks

These codes are also used with the  **Event.GetFilterObjects** and **Event.SetFilterObjects** methods to filter events by object type.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visTypeBitmap**|32|Returned by  **Shape.ForeignType** if the shape is a bitmap.|
| **visTypeDoc**|6|The document's  **DocumentSheet** .|
| **visTypeForeignObject**|4|An imported shape.|
| **visTypeGroup**|2|A shape that contains other shapes.|
| **visTypeGuide**|5|A shape that is a guide.|
| **visTypeInk**|64|Returned by  **Shape.ForeignType** if the shape is ink.|
| **visTypeInval**|0|The type of no shape. Means all types when used as filter code.|
| **visTypeIsControl**|1024|Returned by  **Shape.ForeignType** if the shape is a control.|
| **visTypeIsEmbedded**|512|Returned by  **Shape.ForeignType** if the shape is embedded.|
| **visTypeIsLinked**|256|Returned by  **Shape.ForeignType** if the shape is linked.|
| **visTypeIsOLE2**|32768|Returned by  **Shape.ForeignType** if the shape is linked, embedded, or a control.|
| **visTypeMetafile**|16|Returned by  **Shape.ForeignType** if the shape is a metafile.|
| **visTypePage**|1|Page's or master's  **PageSheet** property.|
| **visTypeShape**|3|Native Microsoft Visio shape.|

