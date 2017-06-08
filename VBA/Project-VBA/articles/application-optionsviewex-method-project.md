---
title: Application.OptionsViewEx Method (Project)
keywords: vbapj.chm646
f1_keywords:
- vbapj.chm646
ms.prod: project-server
api_name:
- Project.Application.OptionsViewEx
ms.assetid: 88abc2b7-116f-4243-f86f-5f4ad9cf8e72
ms.date: 06/08/2017
---


# Application.OptionsViewEx Method (Project)

Sets display options for the  **General**,  **Display**, and  **Advanced** tabs of the **Project Options** dialog box.


## Syntax

 _expression_. **OptionsViewEx**( ** _DefaultView_**, ** _DateFormat_**, ** _ProjectSummary_**, ** _DisplayStatusBar_**, ** _DisplayEntryBar_**, ** _DisplayScrollBars_**, ** _CurrencySymbol_**, ** _SymbolPlacement_**, ** _CurrencyDigits_**, ** _ProjectCurrency_**, ** _DisplayOutlineNumber_**, ** _DisplayOutlineSymbols_**, ** _DisplayNameIndent_**, ** _DisplaySummaryTasks_**, ** _DisplayOLEIndicator_**, ** _DisplayExternalSuccessors_**, ** _DisplayExternalPredecessors_**, ** _CrossProjectLinksInfo_**, ** _AcceptNewExternalData_**, ** _DisplayWindowsInTaskbar_**, ** _DisplayScreentips_**, ** _CalendarType_**, ** _Use3DLook_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DefaultView_|Optional|**String**|The name of the default view.|
| _DateFormat_|Optional|**Long**|The date format. Can be one of the  **[PjDateFormat](pjdateformat-enumeration-project.md)** constants.|
| _ProjectSummary_|Optional|**Boolean**|**True** if the project summary task is visible.|
| _DisplayStatusBar_|Optional|**Boolean**|**True** if the status bar appears.|
| _DisplayEntryBar_|Optional|**Boolean**|**True** if the entry bar appears.|
| _DisplayScrollBars_|Optional|**Boolean**|**True** if scroll bars appear.|
| _CurrencySymbol_|Optional|**String**|The symbol to use for currency values.|
| _SymbolPlacement_|Optional|**Long**|The position to display the currency symbol in currency values. Can be one of the following  **[PjPlacement](pjplacement-enumeration-project.md)** constants: **pjAfter**, **pjAfterWithSpace**, **pjBefore**, or **pjBeforeWithSpace**.|
| _CurrencyDigits_|Optional|**Integer**|The number of digits following the decimal point in currency values.|
| _ProjectCurrency_|Optional|**Variant**|The three-character ISO standard currency code. For example, USD is the code for United States dollars. The  **Currency** drop-down list on the **Display** tab includes all of the currency codes that Project supports.|
| _DisplayOutlineNumber_|Optional|**Boolean**|**True** if the outline numbers for tasks appear.|
| _DisplayOutlineSymbols_|Optional|**Boolean**|**True** if the outline symbols for tasks appear.|
| _DisplayNameIndent_|Optional|**Boolean**|**True** if the names of tasks are indented.|
| _DisplaySummaryTasks_|Optional|**Boolean**|**True** if summary tasks appear.|
| _DisplayOLEIndicator_|Optional|**Boolean**|**True** if the OLE indicator appears.|
| _DisplayExternalSuccessors_|Optional|**Boolean**|**True** if successors in an external project should be displayed.|
| _DisplayExternalPredecessors_|Optional|**Boolean**|**True** if predecessors in an external project should be displayed.|
| _CrossProjectLinksInfo_|Optional|**Boolean**|**True** if the **Links Between Projects** dialog box appears when a project containing cross-project links is opened.|
| _AcceptNewExternalData_|Optional|**Boolean**|**True** if new or changed data from an external project is automatically accepted when a project is opened. If CrossProjectLinksInfo is **True**, AcceptNewExternalData is ignored.|
| _DisplayWindowsInTaskbar_|Optional|**Boolean**|**True** if project windows appear on the task bar and in the task list.|
| _DisplayScreentips_|Optional|**Boolean**|**True** if Project displays screen tips for items such as link lines or column headers.|
| _CalendarType_|Optional|**Integer**|Sets the type of calendar for displaying Project content on the screen. Can be one of the  **[pjCalendarType](pjcalendartype-enumeration-project.md)** values.|
| _Use3DLook_|Optional|**Boolean**|**True** if bars and shapes in Gantt views have a 3-dimensional appearance; otherwise **False**.|

### Return Value

 **Boolean**


## Remarks

If an argument is omitted, the default value is specified by the corresponding setting on the  **General**,  **Display**, or  **Advanced** tab of the **Project Options** dialog box.

Using the  **OptionsViewEx** method without specifying any arguments displays the **Project Options** dialog box with the **General** tab selected.


