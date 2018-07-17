---
title: Worksheet.Change Event (Excel)
keywords: vbaxl10.chm502079
f1_keywords:
- vbaxl10.chm502079
ms.prod: excel
api_name:
- Excel.Worksheet.Change
ms.assetid: d9e11d08-41ba-f0a8-dc55-6c6cd4e76dd0
ms.date: 06/08/2017
---


# Worksheet.Change Event (Excel)

Occurs when cells on the worksheet are changed by the user or by an external link.


## Syntax

 _expression_ . **Change**( **_Target_** )

 _expression_ A variable that represents a **Worksheet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Target_|Required| **[Range](range-object-excel.md)**|The changed range. Can be more than one cell.|

### Return Value

Nothing


## Remarks

This event does not occur when cells change during a recalculation. Use the  **[Calculate](chart-calculate-event-excel.md)** event to trap a sheet recalculation.


## Example

The following code example changes the color of changed cells to blue.


```vb
Private Sub Worksheet_Change(ByVal Target as Range) 
    Target.Font.ColorIndex = 5 
End Sub
```



 **Sample code provided by:** Bill Jelen,[MrExcel.com](http://www.mrexcel.com/) |[About the Contributors](worksheet-change-event-excel.md#AboutContributor)

The following code example verifies that, when a cell value changes, the changed cell is in column A, and if the changed value of the cell is greater than 100. If the value is greater than 100, the adjacent cell in column B is changed to the color red.




```vb
Private Sub Worksheet_Change(ByVal Target As Excel.Range) 
    If Target.Column = 1 Then 
        ThisRow = Target.Row 
        If Target.Value > 100 Then 
            Range("B" &; ThisRow).Interior.ColorIndex = 3 
        Else 
            Range("B" &; ThisRow).Interior.ColorIndex = xlColorIndexNone 
        End If 
    End If 
End Sub
```



 **Sample code provided by:** Tom Urtis,[Atlas Programming Management](http://www.atlaspm.com/) |[About the Contributors](worksheet-change-event-excel.md#AboutContributor)

The following code example sets the values in the range A1:A10 to be uppercase as the data is entered into the cell.




```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Range("A1:A10")) Is Nothing Or Target.Cells.Count > 1 Then Exit Sub
    Application.EnableEvents = False
    'Set the values to be uppercase
    Target.Value = UCase(Target.Value)
    Application.EnableEvents = True
End Sub
```


## About the Contributors
<a name="AboutContributor"> </a>

MVP Bill Jelen is the author of more than two dozen books about Microsoft Excel. He is a regular guest on TechTV with Leo Laporte and is the host of MrExcel.com, which includes more than 300,000 questions and answers about Excel. 

MVP Tom Urtis is the founder of Atlas Programming Management, a full-service Microsoft Office and Excel business solutions company in Silicon Valley. Tom has over 25 years of experience in business management and developing Microsoft Office applications, and is the co author of "Holy Macro! It's 2,500 Excel VBA Examples." 


## See also
<a name="AboutContributor"> </a>


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

