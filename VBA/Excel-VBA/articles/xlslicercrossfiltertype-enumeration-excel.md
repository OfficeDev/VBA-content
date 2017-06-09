---
title: XlSlicerCrossFilterType Enumeration (Excel)
ms.prod: excel
api_name:
- Excel.XlSlicerCrossFilterType
ms.assetid: 8f5e1daa-d548-3e58-4925-07d16c10140d
ms.date: 06/08/2017
---


# XlSlicerCrossFilterType Enumeration (Excel)

Specifies the type of cross filtering used by the specified slicer cache and how it is visualized.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **xlSlicerCrossFilterHideButtonsWithNoData**|4|Cross filtering is turned on for this slicer cache, any tile with no data for a filtering selection in other slicers connected to the same data source will be dimmed. Additionally, buttons will be hidden.|
| **xlSlicerCrossFilterShowItemsWithDataAtTop**|2|Cross filtering is turned on for this slicer cache, any tile with no data for a filtering selection in other slicers connected to the same data source will be dimmed. Additionally, tiles with data are moved to the top in the slicer. (Default)|
| **xlSlicerCrossFilterShowItemsWithNoData**|3|Cross filtering is turned on for this slicer cache, any tile with no data for a filtering selection in other slicers connected to the same data source will be dimmed.|
| **xlSlicerNoCrossFilter**|1|Cross filtering is turned off entirely, so all tiles are displayed and active (not dimmed) regardless of filtering selections in other slicers.|

