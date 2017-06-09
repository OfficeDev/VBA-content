---
title: Application.FilePageSetupView Method (Project)
keywords: vbapj.chm2360
f1_keywords:
- vbapj.chm2360
ms.prod: project-server
api_name:
- Project.Application.FilePageSetupView
ms.assetid: 46a90db8-a635-3592-77ed-c051afa36946
ms.date: 06/08/2017
---


# Application.FilePageSetupView Method (Project)

Sets up view-specific options for printing.


## Syntax

 _expression_. **FilePageSetupView**( ** _Name_**, ** _AllSheetColumns_**, ** _RepeatColumns_**, ** _PrintNotes_**, ** _PrintBlankPages_**, ** _BestPageFitTimescale_**, ** _PrintColumnTotals_**, ** _PrintRowTotals_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the view or report for which to set up pages for printing.|
| _AllSheetColumns_|Optional|**Boolean**|**True** if all table columns print. **False** if only visible table columns print. This argument is only available when the Task Usage view, Resource Usage view, or one of the Gantt views is the active view.|
| _RepeatColumns_|Optional|**Integer**|The number of table columns to print on each page. This argument is only available when the Task Sheet, Task Usage view, Resource Sheet, Resource Usage view, or one of the Gantt views is the active view.|
| _PrintNotes_|Optional|**Boolean**|**True** if notes print. If the active view is the Resource Graph, **PrintNotes** is ignored.|
| _PrintBlankPages_|Optional|**Boolean**|**True** if blank pages print. This argument is only available when the Task Usage view, Resource Usage view, Network Diagram view, or one of the Gantt views is the active view.|
| _BestPageFitTimescale_|Optional|**Boolean**|**True** if the timescale is adjusted so the last printed page is exactly full. This argument is only available when the Task Usage view, Resource Usage view, Resource Graph, or one of the Gantt views is the active view.|
| _PrintColumnTotals_|Optional|**Variant**|NOT available at this time.|
| _PrintRowTotals_|Optional|**Variant**|NOT available at this time.|

### Return Value

 **Boolean**


## Remarks

Using the  **FilePageSetupView** method without specifying any arguments displays the **Page Setup** dialog box with the **View** tab selected.The **FilePageSetupView** method is not available when the Calendar is the active view.


