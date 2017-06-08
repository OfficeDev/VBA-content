---
title: Application.FilePageSetupPage Method (Project)
keywords: vbapj.chm2355
f1_keywords:
- vbapj.chm2355
ms.prod: project-server
api_name:
- Project.Application.FilePageSetupPage
ms.assetid: 7c5cf66d-715b-17e1-a03a-a376617a1e02
ms.date: 06/08/2017
---


# Application.FilePageSetupPage Method (Project)

Sets up pages for printing.


## Syntax

 _expression_. **FilePageSetupPage**( ** _Name_**, ** _Portrait_**, ** _PercentScale_**, ** _PagesTall_**, ** _PagesWide_**, ** _PaperSize_**, ** _FirstPageNumber_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the view or report for which to set up pages for printing.|
| _Portrait_|Optional|**Boolean**|**True** if the page orientation is portrait. **False** if the page orientation is landscape.|
| _PercentScale_|Optional|**Integer**|The scaling factor, specified as a percentage of the original. Can be a number between 1 and 500.|
| _PagesTall_|Optional|**Integer**|The height the printed project should be fit to, in pages. The  **PagesTall** argument is ignored if **PercentScale** is specified.|
| _PagesWide_|Optional|**Variant**| The width the printed project should be fit to, in pages. The **PagesWide** argument is ignored if **PercentScale** is specified.|
| _PaperSize_|Optional|**Long**|The size of the paper to be used when printing. (Some printers may not support all of these paper sizes.) Can be one of the  **[PjPaperSize](pjpapersize-enumeration-project.md)** constants.|
| _FirstPageNumber_|Optional|**String**|Any valid integer to print on the first page or the string "Auto" to print the actual number of the first printed page. (For example, "3" if the first printed page is page 3.) Succeeding page numbers are incremented on this number. The default value is "Auto".|

### Return Value

 **Boolean**


## Remarks

Using the  **FilePageSetupPage** method without specifying any arguments displays the **Page Setup** dialog box with the **Page** tab selected.


