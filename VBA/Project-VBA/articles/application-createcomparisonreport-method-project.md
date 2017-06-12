---
title: Application.CreateComparisonReport Method (Project)
keywords: vbapj.chm2182
f1_keywords:
- vbapj.chm2182
ms.prod: project-server
api_name:
- Project.Application.CreateComparisonReport
ms.assetid: 55b423a7-4613-e1ba-c1b8-e790e74694e7
ms.date: 06/08/2017
---


# Application.CreateComparisonReport Method (Project)

Creates a comparison report between two versions of a project. 


## Syntax

 _expression_. **CreateComparisonReport**( ** _Filename_**, ** _TaskTable_**, ** _ResourceTable_**, ** _Items_**, ** _Columns_**, ** _ShowLegend_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Optional|**String**|Full path and name of the project file to compare.|
| _TaskTable_|Optional|**String**|Name of the table to use for comparison in a task view.|
| _ResourceTable_|Optional|**String**|Name of the table to use for comparison in a resource view.|
| _Items_|Optional|**[PjCompareVersionItems](pjcompareversionitems-enumeration-project.md)**|Specifies the type of items to compare.|
| _Columns_|Optional|**[PjCompareVersionColumns](pjcompareversioncolumns-enumeration-project.md)**|Specifies whether to show only column data, only column differences, or both differences and data.|
| _ShowLegend_|Optional|**Variant**|If  **True**, shows the legend in the comparison report.|

### Return Value

 **Boolean**


## Remarks

The  **CreateComparisonReport** method compares task or resource information, but not assignment information.


## Example

The following example demonstrate how to create a comparison report. The code first checks that a project is currently open, and then checks that either tasks or resources are in the project. The comparison report is based on cost tables, filtered for only changed task or resource cost information, with columns that display only the differences between tasks or resources. Finally, the comparison report is saved with a file name based on the current (first) project.


```vb
Sub ComparisonReport () 
    If Projects.Count = 0 Then 
        MsgBox "You must have at least one active project open before you can compare projects.", _ 
            vbInformation 
        Exit Sub 
    ElseIf ActiveProject.Tasks.Count = 0 Then 
        If ActiveProject.ResourceCount = 0 Then 
            MsgBox "There are no task or resources in the current project. " &; vbCrLf _ 
            &; "Open a project with either tasks or resources before creating a comparison report.", _ 
            vbInformation 
            Exit Sub 
        End If 
    End If 
 
    ' Get the name of the project to use for saving the comparison report. 
    Dim currentProject As Project 
    Set currentProject = ActiveProject 
 
    Dim previousVersion As String 
    previousVersion = "[full path to .mpp file to compare with the active project.]" 
 
    CreateComparisonReport FileName:=previousVersion, _ 
    TaskTable:="Cost", _ 
    ResourceTable:="Cost", _ 
    Items:=pjCompareVersionItemsChangedItems, _ 
    Columns:=pjCompareVersionColumnsDifferencesOnly, _ 
    Showlegend:=True 
 
    ' Save the comparison report based upon the name of the first project. 
    Dim comparisonReport As Project 
    Set comparisonReport = ActiveProject 
    ActiveProject.SaveAs currentProject &; "_Compared.mpp" 
End Sub
```


