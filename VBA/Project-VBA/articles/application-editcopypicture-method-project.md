---
title: Application.EditCopyPicture Method (Project)
keywords: vbapj.chm204
f1_keywords:
- vbapj.chm204
ms.prod: project-server
api_name:
- Project.Application.EditCopyPicture
ms.assetid: 03f6306b-3538-9a34-dbc3-4ff2f7f40b1e
ms.date: 06/08/2017
---


# Application.EditCopyPicture Method (Project)

Copies the active view as a picture or an OLE object, or exports the active view to a GIF image file.


## Syntax

 _expression_. **EditCopyPicture**( ** _Object_**, ** _ForPrinter_**, ** _SelectedRows_**, ** _FromDate_**, ** _ToDate_**, ** _Filename_**, ** _ScaleOption_**, ** _MaxImageHeight_**, ** _MaxImageWidth_**, ** _MeasurementUnits_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Object_|Optional|**Boolean**|**True** if the view should be copied as an OLE object. The default value is **False**.|
| _ForPrinter_|Optional|**Long**|Specifies where to copy the view. Can be one of the following  **PjCopyPictureFor** constants: **pjScreen**, **pjPrinter**, or **pjGIF**. If **Object** is **True**, **ForPrinter** is ignored. The default value is **pjScreen**.|
| _SelectedRows_|Optional|**Boolean**|**True** if Project copies only the selected rows. **False** if the program copies all visible rows.|
| _FromDate_|Optional|**Variant**|The beginning of the timescale for the copied picture. If  **Object** is **True**, **FromDate** is ignored. If **FromDate** is specified and **ToDate** is not, Project will use the last entered date for the end of the timescale. If that would create a negative time span, the program will use the latest timescale date visible in the active view. The default value is the earliest timescale date visible in the active view.|
| _ToDate_|Optional|**Variant**|The end of the timescale for the copied picture. If  **Object** is **True**, **ToDate** is ignored. If **ToDate** is specified and **FromDate** is not, Project will use the last entered date for the beginning of the timescale. If that would create a negative time span, the program will use the earliest timescale date visible in the active view. The default value is the latest timescale date visible in the active view.|
| _Filename_|Optional|**String**|The file name for the GIF image file. If  **ForPrinter** is **pjGIF**, **FileName** is required. If **Object** is **True**, or **ForPrinter** is not **pjGIF**, **FileName** is ignored.|
| _ScaleOption_|Optional|**Long**|Specifies how to treat a picture of the active view if it is larger than  **MaxImageWidth** by **MaxImageHeight** (default 22 inches by 22 inches). The default value is **pjCopyPictureKeepRange**. Can be one of the **[PjCopyPictureScaleOption](pjcopypicturescaleoption-enumeration-project.md)** constants.|
| _MaxImageHeight_|Optional|**Double**|Specifies the maximum height of the picture. The accepted range of  **MaxImageHeight** is 1 to 22 inches (2.54 to 55.88 centimeters). The default value is 22 (inches).|
| _MaxImageWidth_|Optional|**Double**|Specifies the maximum width of the picture. The accepted range of  **MaxImageWidth** is 1 to 22 inches (2.54 to 55.88 centimeters). The default value is 22 (inches).|
| _MeasurementUnits_|Optional|**Variant**|**Long**. Specifies the units of measurement. The default value is **pjInches**. Can be one of the **[PjMeasurementUnits](pjmeasurementunits-enumeration-project.md)** constants.|

### Return Value

 **Boolean**


## Remarks

Using the  **EditCopyPicture** method with no arguments displays the **Copy Picture** dialog box.


## Example

The following example makes a copy of the Gantt Chart view as Test.gif file and saves in the root folder.


```vb
Sub Edit_CopyPicture() 
    'Activate the Gantt Chart view 
    ViewApply Name:="&;Gantt Chart" 
    EditCopyPicture ForPrinter:=pjGIF, FileName:="C:\Test.gif" 
End Sub
```


