---
title: Shapes Object (Project)
ms.prod: project-server
ms.assetid: 6e42040c-dd5a-de4c-afa8-f9e33d1e5054
ms.date: 06/08/2017
---


# Shapes Object (Project)
Represents a collection of  **[Shape](http://msdn.microsoft.com/library/d2b32bcd-5595-a4a7-9772-feb25fd0103a%28Office.15%29.aspx)** objects in a custom report.

## Example

Use the  **[Report.Shapes](http://msdn.microsoft.com/library/2f62c406-3845-79f8-3d17-e5891c1e23f9%28Office.15%29.aspx)** property to get the **Shapes** collection object. In the following example, the report must be the active view to get the **Shapes** collection; otherwise, you get a run-time error 424 (Object required) in the `For Each oShape In oReport.Shapes` statement.


```
Sub ListShapesInReport()
    Dim oReports As Reports
    Dim oReport As Report
    Dim oShape As shape
    Dim reportName As String
    Dim msg As String
    Dim msgBoxTitle As String
    Dim numShapes As Integer
    
    numShapes = 0
    msg = ""
    reportName = "Table Tests"
    Set oReports = ActiveProject.Reports
    
    If oReports.IsPresent(reportName) Then
        ' Make the report the active view.
        oReports(reportName).Apply
        
        Set oReport = oReports(reportName)
        msgBoxTitle = "Shapes in report: '" &amp; oReport.Name &amp; "'"
    
        For Each oShape In oReport.Shapes
            numShapes = numShapes + 1
            msg = msg &amp; numShapes &amp; ". Shape type: " &amp; CStr(oShape.Type) _
                &amp; ", '" &amp; oShape.Name &amp; "'" &amp; vbCrLf
        Next oShape
        
        If numShapes > 0 Then
            MsgBox Prompt:=msg, Title:=msgBoxTitle
        Else
            MsgBox Prompt:="This report contains no shapes.", _
                Title:=msgBoxTitle
        End If
    Else
         MsgBox Prompt:="The requested report, '" &amp; reportName _
            &amp; "', does not exist.", Title:="Report error"
    End If
End Sub
```


## Methods



|**Name**|
|:-----|
|[AddCallout](http://msdn.microsoft.com/library/6c183677-d869-f493-7226-14cca4329aae%28Office.15%29.aspx)|
|[AddChart](http://msdn.microsoft.com/library/d404a9de-c1aa-c2a0-bf85-dc1f1735cf3c%28Office.15%29.aspx)|
|[AddConnector](http://msdn.microsoft.com/library/bfd75cf3-f70b-8d19-bf28-94e2f4b227dd%28Office.15%29.aspx)|
|[AddCurve](http://msdn.microsoft.com/library/16ea0f55-268a-b224-cc94-3d7e74de6265%28Office.15%29.aspx)|
|[AddLabel](http://msdn.microsoft.com/library/3fd21dbc-51b7-0e22-8c8a-359b1717932f%28Office.15%29.aspx)|
|[AddLine](http://msdn.microsoft.com/library/697a5972-4b24-8e77-b42f-b064019906fa%28Office.15%29.aspx)|
|[AddPolyline](http://msdn.microsoft.com/library/c61cbaf3-b687-b137-e4a2-8f9061dfc0f0%28Office.15%29.aspx)|
|[AddShape](http://msdn.microsoft.com/library/58af0a51-a455-5c9a-1cae-e56dc67a08a5%28Office.15%29.aspx)|
|[AddTable](http://msdn.microsoft.com/library/d4f9942b-ebd5-20e6-c8d4-f7107d1e1eab%28Office.15%29.aspx)|
|[AddTextbox](http://msdn.microsoft.com/library/ee8c619f-8b35-6f94-e680-86dbeedd6d19%28Office.15%29.aspx)|
|[AddTextEffect](http://msdn.microsoft.com/library/5510367c-7f8d-3266-642f-61f3d45a18cf%28Office.15%29.aspx)|
|[BuildFreeform](http://msdn.microsoft.com/library/257f76e3-3b37-5b58-cb78-f6fcebe1ca29%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/43fba4f4-f3d3-20a0-2c77-15e31dcdcbf5%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/984326ae-f567-18b8-562a-fcb2160b0dad%28Office.15%29.aspx)|
|[SelectAll](http://msdn.microsoft.com/library/f85eb8ea-770f-ba13-b7d4-794d162bd598%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Background](http://msdn.microsoft.com/library/9199c72e-d692-6a9c-2ff2-06fe9e445bef%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/c198cf75-b554-5815-4b77-d2a54d60f5e6%28Office.15%29.aspx)|
|[Default](http://msdn.microsoft.com/library/46895c7b-6cb1-0286-1e9d-8cc658ea6441%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/ca0ec6c1-657d-517b-eebe-6a5b20bbe21f%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/f10fef14-baee-ddd3-fb39-81fef0bc132d%28Office.15%29.aspx)|

## See also


#### Other resources


[Shape Object](http://msdn.microsoft.com/library/d2b32bcd-5595-a4a7-9772-feb25fd0103a%28Office.15%29.aspx)
[Report Object](http://msdn.microsoft.com/library/38ef993e-e5cd-b451-06aa-41eb0e93450e%28Office.15%29.aspx)
[ShapeRange Object](http://msdn.microsoft.com/library/315031aa-4b8c-424b-26e7-ce15897beb05%28Office.15%29.aspx)
