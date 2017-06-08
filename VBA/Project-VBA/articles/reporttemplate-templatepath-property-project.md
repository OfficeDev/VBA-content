---
title: ReportTemplate.TemplatePath Property (Project)
ms.prod: project-server
api_name:
- Project.ReportTemplate.TemplatePath
ms.assetid: be8381a8-f19e-76f0-32c8-c85f29ba93cc
ms.date: 06/08/2017
---


# ReportTemplate.TemplatePath Property (Project)

Gets the path and file name of the Visual Report template. Read-only  **String**.


## Syntax

 _expression_. **TemplatePath**

 _expression_ A variable that represents a **ReportTemplate** object.


## Remarks

The Visual Report template files are stored in the following directory for each user , where LCID is the language code identifier such as 1033 for U.S. English:  `C:\Users\[UserAlias]\AppData\Roaming\Microsoft\Templates\[LCID]\`. For example, adr1.xlt is a Microsoft Excel template.


## Example

The following example lists all of the Visual Report template types and files for the current user.


```vb
Sub ListTemplatePaths() 

 Dim templateList As String 

 Dim typeOfTemplate As String 

 Dim template As ReportTemplate 

 

 For Each template In Application.VisualReportTemplateList 

 Select Case template.templateType 

 Case pjExcel 

 typeOfTemplate = "Excel" 

 Case pjVisioMetric 

 typeOfTemplate = "Visio Metric" 

 Case pjVisioUS 

 typeOfTemplate = "Visio U.S." 

 Case Else 

 End Select 

 

 templateList = templateList &; vbCrLf &; typeOfTemplate &; ": " _ 

 &; template.TemplatePath 

 Next template 

 

 MsgBox "Visual Reports Templates:" &; templateList 

 

End Sub
```


## See also


#### Concepts


[ReportTemplate Object](reporttemplate-object-project.md)
