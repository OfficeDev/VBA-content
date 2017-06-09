---
title: Application.PasteSourceFormatting Method (Project)
keywords: vbapj.chm139
f1_keywords:
- vbapj.chm139
ms.prod: project-server
ms.assetid: 3544cad7-51d4-fd80-5aaa-396fb26a0d17
ms.date: 06/08/2017
---


# Application.PasteSourceFormatting Method (Project)
Pastes a copy of a report or a shape, where the copy maintains the formatting of the source.

## Syntax

 _expression_. **PasteSourceFormatting**

 _expression_ A variable that represents an **Application** object.


### Return value

 **Boolean**

 **True** if the paste is successful; otherwise, **False**.


## Example

The following example copies the built-in Cost Report, creates a custom report, pastes the copied report into the new report by using the source formatting, and then renames the report title.


```vb
Sub CopyCostReport()
    Dim reportName As String
    Dim newReportName As String
    Dim newReportTitle As String
    Dim myNewReport As Report
    Dim oShape As Shape
    Dim msg As String
    Dim msgBoxTitle As String
    Dim numShapes As Integer
    
    reportName = "Task Cost Overview"   ' Built-in report
    newReportName = "Task Cost Copy 2"
    msg = ""
    numShapes = 0
    
    If ActiveProject.Reports.IsPresent(reportName) Then
        ApplyReport reportName
        CopyReport
        Set myNewReport = ActiveProject.Reports.Add(newReportName)
        PasteSourceFormatting
        
        ' List the shapes in the copied report.
        For Each oShape In myNewReport.Shapes
            numShapes = numShapes + 1
            msg = msg &; numShapes &; ". Shape type: " &; CStr(oShape.Type) _
                &; ", '" &; oShape.Name &; "'" &; vbCrLf
            
            ' Modify the report title.
            If oShape.Name = "TextBox 1" Then
                newReportTitle = "My " &; oShape.TextFrame2.TextRange.Text
                With oShape.TextFrame2.TextRange
                    .Text = newReportTitle
                    .Characters.Font.Fill.ForeColor.RGB = &;H60FF10 ' Bluish green.
                End With
                
                oShape.Reflection.Type = msoReflectionType2
                oShape.IncrementTop -10    ' Move the title 10 points up.
                oShape.Select
            End If
        Next oShape
        
        msgBoxTitle = "Shapes in report: '" &; myNewReport.Name &; "'"
                
        If numShapes > 0 Then
            MsgBox Prompt:=msg, Title:=msgBoxTitle
        Else
            MsgBox Prompt:="This report contains no shapes.", _
                Title:=msgBoxTitle
        End If
    Else
        MsgBox Prompt:="No custom report name: " &; reportName, _
            Title:="ApplyReport error", Buttons:=vbExclamation
    End If
End Sub
```


## See also


#### Concepts


[Application Object](application-object-project.md)
#### Other resources


[CopyReport Method](application-copyreport-method-project.md)
[Shape.Copy Method](shape-copy-method-project.md)
[PasteDestFormatting Method](application-pastedestformatting-method-project.md)
[PasteAsPicture Method](application-pasteaspicture-method-project.md)
