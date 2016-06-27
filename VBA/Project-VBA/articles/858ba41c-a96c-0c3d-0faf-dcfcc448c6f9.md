
# Chart.DataTable Property (Project)
Gets an  **Office.IMsoDataTable** object that represents the chart data table. Read-only **IMsoDataTable**.

## Syntax

 _expression_. **DataTable**

 _expression_ A variable that represents a **Chart** object.


## Remarks

To see the  **IMsoDataTable** object, right-click in the Object Browser, and then choose **Show Hidden Members**.


## Example

The following example adds a data table with an outline border to the chart on the active report.


```vb
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

 **IMSODATATABLE**


## See also


#### Other resources


[Chart Object](810d4ec1-69d2-c432-b9da-57042b783b85.md)