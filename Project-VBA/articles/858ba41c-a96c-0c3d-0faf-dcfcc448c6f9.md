
# Chart.DataTable Property (Project)
Gets an  **Office.IMsoDataTable** object that represents the chart data table. Read-only **IMsoDataTable**.

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)
 [Property value](#sectionSection3)


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **DataTable**

 _expression_A variable that represents a  **Chart** object.


## Remarks
<a name="sectionSection1"> </a>

To see the  **IMsoDataTable** object, right-click in the Object Browser, and then choose **Show Hidden Members**.


## Example
<a name="sectionSection2"> </a>

The following example adds a data table with an outline border to the chart on the active report.


```
Sub ShowDataTable()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    With chartShape.Chart
        .HasDataTable = True
        .DataTable.HasBorderOutline = True
    End With
End Sub
```


## Property value
<a name="sectionSection3"> </a>

 **IMSODATATABLE**


## See also
<a name="sectionSection3"> </a>


#### Other resources


 [Chart Object](810d4ec1-69d2-c432-b9da-57042b783b85.md)
