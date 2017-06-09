---
title: Application.FilePageSetupCalendarTextEx Method (Project)
keywords: vbapj.chm2162
f1_keywords:
- vbapj.chm2162
ms.prod: project-server
api_name:
- Project.Application.FilePageSetupCalendarTextEx
ms.assetid: 370cfaa4-4a7b-e40e-be9e-d562bf9947d7
ms.date: 06/08/2017
---


# Application.FilePageSetupCalendarTextEx Method (Project)

Formats the text of calendar views for printing, where the text color can be a hexadecimal value.


## Syntax

 _expression_. **FilePageSetupCalendarTextEx**( ** _Name_**, ** _Item_**, ** _Font_**, ** _Size_**, ** _Bold_**, ** _Italic_**, ** _Underline_**, ** _Color_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the calendar to edit.|
| _Item_|Optional|**Long**|The text item to format. Can be one of the  **[PjPageSetupCalendarItem](pjpagesetupcalendaritem-enumeration-project.md)** constants.|
| _Font_|Optional|**String**|The name of the font.|
| _Size_|Optional|**Integer**|The size of the font in points|
| _Bold_|Optional|**Boolean**|**True** if the font is bold; otherwise, **False**.|
| _Italic_|Optional|**Boolean**|**True** if the font is italic; otherwise, **False**.|
| _Underline_|Optional|**Boolean**|**True** if the font is underlined; otherwise, **False**.|
| _Color_|Optional|**Long**|The color of the text. Can be a hexadecimal RGB value, where red is the last byte. For example, the value &;H01FFFF is yellow.|

### Return Value

 **Boolean**


## Remarks

Using the  **FilePageSetupCalendarTextEx** method without any arguments displays the **Text Styles** dialog box.


 **Note**   **FilePageSetupCalendarTextEx** works only for printing calendar views.


## Example

The following example formats monthly titles in red for printing.


```vb
Sub File_PageSetupCalendarText() 
 
    'Activate the Calandar view. 
    ViewApply Name:="&;Calendar" 
 
    FilePageSetupCalendarTextEx Item:=pjMonthlyTitles, Color:=&;0101FF 
    FilePrint 
End Sub
```


