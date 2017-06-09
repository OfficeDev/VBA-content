---
title: Application.UsageViewEntryEx Method (Project)
keywords: vbapj.chm2163
f1_keywords:
- vbapj.chm2163
ms.prod: project-server
api_name:
- Project.Application.UsageViewEntryEx
ms.assetid: 2aac9824-ab5c-006d-99d2-07e019e6409d
ms.date: 06/08/2017
---


# Application.UsageViewEntryEx Method (Project)

Adds fields to the  **Details** pane and option menu for the Task Usage or Resource Usage views, and formats the styles to help distinguish detail rows.


## Syntax

 _expression_. **UsageViewEntryEx**( ** _CurIndex_**, ** _Order_**, ** _FontWord_**, ** _CellBackground_**, ** _Pattern_**, ** _Shortcut_**, ** _DisplayField_**, ** _FontColor_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CurIndex_|Optional|**Integer**|Current zero-based index of fields in the  **Show these fields** list in the **Detail Styles** dialog box. Values greater than the number of fields currently shown are reduced to the next highest value in the actual list. For example, if there are two fields showing and _CurIndex_ = 8, the value of _CurIndex_ is reduced to 2. The default value is 0.|
| _Order_|Optional|**Integer**|Order of the field in an internal array of fields. For valid values, see the table of field names in the Remarks section.|
| _FontWord_|Optional|**Long**|Deprecated in Project. In some versions of Project,  _FontWord_ set the font color by using the **PjColor** enumeration.|
| _CellBackground_|Optional|**Long**|Color of the cells for entry. Can be a hexadecimal RGB value, where red is the last byte. For example, &;HFFFF00 is blue-green.|
| _Pattern_|Optional|**Integer**|Background pattern of the cells for entry. Can be one of the  **[PjBackgroundPattern](pjbackgroundpattern-enumeration-project.md)** constants.|
| _Shortcut_|Optional|**Boolean**|**True** if the field is shown on the option menu of the **Details** pane; otherwise, **False**. Shortcut is **True** if DisplayField is **True**.|
| _DisplayField_|Optional|**Boolean**|**True** if the field is displayed in the **Details** pane; otherwise, **False**. The DisplayField parameter has no effect on fields that are already displayed.|
| _FontColor_|Optional|**Long**|Color of text in the  **Details** column for usage entry. Can be a hexadecimal RGB value, where red is the last byte. For example, &;HFF00FF is purple.|

### Return Value

 **Boolean**


## Remarks

In the Task Usage or Resource Usage view, choose the  **FORMAT** tab to see the six default fields in the **Details** group on the ribbon. The **Add Details** command displays the **Detail Styles** dialog box, which shows?in alphabetical order?all of the fields available in the current view.

The following table lists the possible fields for the  _Order_ parameter, and shows values for the Task Usage and Resource Usage views.


||||
|:-----|:-----|:-----|
|**Field name for the  _Order_ parameter**|**Task Usage Value**|**Resource Usage Value**|
|Work|0|0|
|Overtime Work|1|1|
|Actual Work|2|2|
|Actual Overtime Work|3|3|
|Cumulative Work |4|4|
|Baseline Work |5|5|
|Overallocation |6|6|
|Percent Allocation |7|7|
|Peak Units |8|8|
|Cost |9|9|
|Fixed Cost |10|N/A|
|Actual Cost |11|10|
|Baseline Cost |12|11|
|Cumulative Cost |13|12|
|BCWS |14|13|
|BCWP |15|14|
|ACWP |16|15|
|SV|17|16|
|CV|18|17|
|Regular Work |19|18|
|Remaining Availability|N/A|19|
|Unit Availability|N/A|20|
|Work Availability|N/A|21|
|Percent Complete |20|N/A|
|Cumulative Percent Complete|21|N/A|
|Baseline 1 Work |22|22|
|Baseline 1 Cost |23|23|
|Baseline 2 Work |24|24|
|Baseline 2 Cost|25|25|
|Baseline 3 Work|26|26|
|Baseline 3 Cost|27|27|
|Baseline 4 Work|28|28|
|Baseline 4 Cost|29|29|
|Baseline 5 Work|30|30|
|Baseline 5 Cost|31|31|
|Baseline 6 Work|32|32|
|Baseline 6 Cost|33|33|
|Baseline 7 Work|34|34|
|Baseline 7 Cost|35|35|
|Baseline 8 Work|36|36|
|Baseline 8 Cost|37|37|
|Baseline 9 Work|38|38|
|Baseline 9 Cost|39|39|
|Baseline 10 Work|40|40|
|Baseline 10 Cost|41|41|
|Actual Fixed Cost |42|N/A|
|CPI |43|N/A|
|SPI |44|N/A|
|CV Percent|45|N/A|
|SV Percent|46|N/A|
|Budget Work |47|42|
|Budget Cost |48|43|
|Baseline Budget Work|49|44|
|Baseline Budget Cost|50|45|
|Baseline 1 Budget Work|51|46|
|Baseline 1 Budget Cost|52|47|
|Baseline 2 Budget Work|53|48|
|Baseline 2 Budget Cost|54|49|
|Baseline 3 Budget Work|55|50|
|Baseline 3 Budget Cost|56|51|
|Baseline 4 Budget Work|57|52|
|Baseline 4 Budget Cost|58|53|
|Baseline 5 Budget Work|59|54|
|Baseline 5 Budget Cost|60|55|
|Baseline 6 Budget Work|61|56|
|Baseline 6 Budget Cost|62|57|
|Baseline 7 Budget Work|63|58|
|Baseline 7 Budget Cost|64|59|
|Baseline 8 Budget Work|65|60|
|Baseline 8 Budget Cost|66|61|
|Baseline 9 Budget Work|67|62|
|Baseline 9 Budget Cost|68|63|
|Baseline 10 Budget Work|69|64|
|Baseline 10 Budget Cost|70|65|
|All Task Rows |71|N/A|
|All Resource Rows|N/A|66|
|All Assignment Rows |72|67|
 In Project 2003 and Office Project 2007, the original **UsageViewEntry** method was not exposed in the VBA object model.


## Example

In the Resource Usage view, the following statement colors the cells for data entry a light yellow in a diagonal-left pattern and colors the  **Work** text in the **Details** column purple to help show the rows for data entry.


```vb
Application.UsageViewEntryEx CellBackground:=&;H01ffff, Pattern:=pjBackgroundDiagonalLeft, _ 
 FontColor:=&;Hff00ff
```

In the Task Usage view, the default field is  **Work**. The following statement adds the  **Actual Cost** field in green, after the **Work** field.




```vb
Application.UsageViewEntryEx Order:=11, CurIndex:=1, DisplayField:=True, FontColor:=&;H10FF10
```


