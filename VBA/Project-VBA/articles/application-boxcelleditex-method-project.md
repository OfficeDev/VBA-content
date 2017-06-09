---
title: Application.BoxCellEditEx Method (Project)
keywords: vbapj.chm2156
f1_keywords:
- vbapj.chm2156
ms.prod: project-server
api_name:
- Project.Application.BoxCellEditEx
ms.assetid: 86405780-ea5f-d32b-b2e5-3d3999c1877d
ms.date: 06/08/2017
---


# Application.BoxCellEditEx Method (Project)

Sets the properties of an individual cell in a data template for a Network Diagram view (PERT chart), where the font color can be specified with a hexadecimal value.


## Syntax

 _expression_. **BoxCellEditEx**( ** _Name_**, ** _Cell_**, ** _FieldName_**, ** _Font_**, ** _FontSize_**, ** _FontColor_**, ** _Bold_**, ** _Italic_**, ** _Underline_**, ** _HorizontalAlignment_**, ** _VerticalAlignment_**, ** _TextLineLimit_**, ** _ShowLabel_**, ** _Label_**, ** _DateFormat_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**| The name of the data template containing the cell to edit.|
| _Cell_|Required|**Long**|The cell to edit. Can be one of the  **[PjCell](pjcell-enumeration-project.md)** constants.|
| _FieldName_|Optional|**Long**|The name of the field to display in the cell. Can be one of the  **[PjField](pjfield-enumeration-project.md)** constants.|
| _Font_|Optional|**String**|The name of the font.|
| _FontSize_|Optional|**Integer**|The size of the font, in points.|
| _FontColor_|Optional|**Long**|The RGB color of the font. Can be a hexadecimal value, where red is the last byte. For example, the value &;HFF0000 is blue and &;H00FFFF is yellow.|
| _Bold_|Optional|**Boolean**|**True** if the font is bold; otherwise, **False**.|
| _Italic_|Optional|**Boolean**|**True** if the font is italic; otherwise, **False**.|
| _Underline_|Optional|**Boolean**|**True** if the font is underlined; otherwise, **False**.|
| _HorizontalAlignment_|Optional|**Long**|Specifies the horizontal alignment of text in the cell. Can be one of the  **[PjAlignment](pjalignment-enumeration-project.md)** constants.|
| _VerticalAlignment_|Optional|**Long**|Specifies the horizontal alignment of text in the cell. Can be one of the  **[PjVerticalAlignment](pjverticalalignment-enumeration-project.md)** constants.|
| _TextLineLimit_|Optional|**Long**|Specifies the limit for the number of lines of text in the cell. Values can be 1, 2, or 3. |
| _ShowLabel_|Optional|**Boolean**|**True** if a label is shown in the cell; otherwise, **False**.|
| _Label_|Optional|**String**|Specifies the cell label.|
| _DateFormat_|Optional|**Long**|Specifies the date format for the cell when  **FieldName** is a date field. Can be one of the **[PjDateFormat](pjdateformat-enumeration-project.md)** constants.|

### Return Value

 **Boolean**


## Remarks

Using the  **BoxCellEditEx** method with only the **Name** and **Cell** arguments has no effect.


## Example

The following example modifies a copy of the  **Critical** data template named **Test Critical**. The macro removes the fourth row of cells and sets the fourth cell in the third row to show the  **Actual Cost** field and label in a purple-blue color.


```vb
Sub ModifyCriticalDataTemplate() 
 Application.BoxCellLayout Name:="Test Critical", CellRows:=3, MergeCells:=True 
 
 Application.BoxCellEditEx Name:="Test Critical", Cell:=pjCell4_3, _ 
 FieldName:=PjField.pjTaskActualCost, Font:="Arial", FontSize:="8", FontColor:=&;HFF0077, _ 
 Bold:=False, Italic:=False, Underline:=False, HorizontalAlignment:=pjLeft, _ 
 VerticalAlignment:=pjMiddle, TextLineLimit:=1, ShowLabel:=True, Label:="Cost" 
End Sub
```


 **Note**  If you use any of the  **PjColor** enumeration constants for the _FontColor_ parameter, the color will be nearly black. For example, the value of **pjGreen** is 9, which in the **BoxCellEditEx** method is a very dark red. To use only the sixteen colors available with **PjColor** constants, use the **[BoxCellEdit](application-boxcelledit-method-project.md)** method.


