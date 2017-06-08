---
title: Application.ReportPrint Method (Project)
keywords: vbapj.chm110
f1_keywords:
- vbapj.chm110
ms.prod: project-server
api_name:
- Project.Application.ReportPrint
ms.assetid: 4117b555-2985-f129-65aa-9f6804ebf221
ms.date: 06/08/2017
---


# Application.ReportPrint Method (Project)

Deprecated in Project. Prints a report.


## Syntax

 _expression_. **ReportPrint**( ** _Name_**, ** _FromPage_**, ** _ToPage_**, ** _PageBreaks_**, ** _Draft_**, ** _Copies_**, ** _FromDate_**, ** _ToDate_**, ** _Preview_**, ** _Color_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the report to print|
| _FromPage_|Optional|**Integer**|A number that specifies the first page to print.|
| _ToPage_|Optional|**Integer**|A number that specifies the last page to print.|
| _PageBreaks_|Optional|**Boolean**|**True** if Project uses manual page breaks when printing. The default value is **True**.|
| _Draft_|Optional|**Boolean**|**True** if Project prints the report in draft mode. The default value is **False**.|
| _Copies_|Optional|**Integer**|A number that specifies the number of copies to print. The default value is 1.|
| _FromDate_|Optional|**Variant**|A number or string that specifies the first date to print. The default value is the start date of the project.|
| _ToDate_|Optional|**Variant**|A number or string that specifies the last date to print. The default value is the finish date of the project.|
| _Preview_|Optional|**Boolean**|**True** if Project previews the active view rather than printing it. The default value is **False**.|
| _Color_|Optional|**Boolean**|**True** if Project prints the report in color. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

In Project, the  **ReportPrint** method returns error 1100, "The method is not available in this situation." In Project, using the **ReportPrint** method no arguments displays the **Custom Reports** dialog box.


## Example

The following example creates a consolidated project, prints a report, and closes the consolidated project without saving it.


```vb
Sub ConsolidatedReport() 
    ConsolidateProjects Filenames:="project1.mpp" &; ListSeparator &; "project2.mpp" 
    ReportPrint Name:="Project Summary" 
    FileClose Save:=pjDoNotSave 
End Sub
```


