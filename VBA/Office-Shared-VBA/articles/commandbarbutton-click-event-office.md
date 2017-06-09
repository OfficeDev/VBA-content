---
title: CommandBarButton.Click Event (Office)
keywords: vbaof11.chm230001
f1_keywords:
- vbaof11.chm230001
ms.prod: office
api_name:
- Office.CommandBarButton.Click
ms.assetid: d4f970e6-8c37-c5cc-a0b4-4efe213a2e05
ms.date: 06/08/2017
---


# CommandBarButton.Click Event (Office)

Occurs when the user clicks a  **CommandBarButton** object.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Click**( **_Ctrl_**, **_CancelDefault_** )

 _expression_ A variable that represents a **CommandBarButton** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Ctrl_|Required|**CommandBarButton**|Represents a CommandBar button|
| _CancelDefault_|Required|**Boolean**|Is  **False** if the default behavior associated with the CommandBarButton controls occurs, unless it's canceled by another process or add-in.|

## Remarks

The  **Click** event is recognized by the **CommandBarButton** object. To return the **Click** event for a particular **CommandBarButton** control, use the **WithEvents** keyword to declare a variable, and then set the variable to the control.


## Example

The following example creates a command bar button on the  **File** menu of the host application that enables the user to save a workbook as a comma-separated value file. (This example works in all applications, but the context of saving as CSV is applicable to Microsoft Excel.)


```
Private HostApp As Object 
 
Sub createAndSynch() 
    Dim iIndex As Integer 
    Dim iCount As Integer 
    Dim fBtnExists As Boolean 
     
    Dim obCmdBtn As Object 
    Dim btnSaveAsCSVHandler as new Class1 
          
    Set HostApp = Application 
     
    Dim barHelp As Office.CommandBar 
    Set barHelp = Application.CommandBars("File") 
    fBtnExists = False  
    iCount = barHelp.Controls.Count 
    For iIndex = 1 To iCount 
        If barHelp.Controls(iIndex).Caption = "Save As CSV (Comma Delimited)" Then fBtnExists = True  
     
    Next 
    Dim btnSaveAsCSV As Office.CommandBarButton 
    If fBtnExists Then 
        Set btnSaveAsCSV = barHelp.Controls("Save As CSV (Comma Delimited)") 
    Else 
        Set btnSaveAsCSV = barHelp.Controls.Add(msoControlButton) 
        btnSaveAsCSV.Caption = "Save As CSV (Comma Delimited)" 
    End If 
     
    btnSaveAsCSV.Tag = "btn1" 
    btnSaveAsCSVHandler.SyncButton btnSaveAsCSV 
End Sub
```


## See also


#### Concepts


[CommandBarButton Object](commandbarbutton-object-office.md)
#### Other resources


[CommandBarButton Object Members](commandbarbutton-members-office.md)

