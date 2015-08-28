
# Shapes.AddChart Method (Project)
Creates a chart at the specified location on the active report. Returns a  **Shape** object that represents the chart.

 **Last modified:** July 28, 2015


## Syntax

 _expression_. **AddChart**(Style,Type,Left,Top,Width,Height,NewLayout)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Style|Optional| **Integer**|Specifies the color style of the chart. The values correspond to the  **Change Colors** drop-down list on the **Chart Styles** group, on the **DESIGN** tab, under **CHART TOOLS** on the ribbon (but the values are not in the same order).|
|Type|Optional| **XlChartType**|The type of chart to add, such as a column chart or pie chart.|
|Left|Optional| **Single**|The position, measured in points, of the left edge of the chart.|
|Top|Optional| **Single**|The position, measured in points, of the top edge of the chart.|
|Width|Optional| **Single**|The width of the chart, measured in points.|
|Height|Optional| **Single**|The height of the chart, measured in points.|
|NewLayout|Optional| **Boolean**|NewLayout is not used in Project.|
|Style|Optional|INT||
|Type|Optional|XLCHARTTYPE||
|Left|Optional|FLOAT||
|Top|Optional|FLOAT||
|Width|Optional|FLOAT||
|Height|Optional|FLOAT||
|NewLayout|Optional|BOOL||

### Return value

 **Shape**


## Example

The following example creates a report that has a default bar chart type with orange-colored bars.


```
Sub AddDefaultChart()
    Dim chartReport As Report
    Dim reportName As String
    
    ' Add a report.
    reportName = "Test chart report"
    Set chartReport = ActiveProject.Reports.Add(reportName)

    ' Add a chart.
    Dim chartShape As shape
    Set chartShape = ActiveProject.Reports(reportName).Shapes.AddChart(Style:=12)
    
    With chartShape
        .Chart.SetElement msoElementChartTitleAboveChart
        .Chart.ChartTitle.Text = "Test Chart"
    End With
End Sub
```


## See also


#### Other resources


 [Shapes Object](6e42040c-dd5a-de4c-afa8-f9e33d1e5054.md)
 [Shape Object](d2b32bcd-5595-a4a7-9772-feb25fd0103a.md)
 [Chart Object](810d4ec1-69d2-c432-b9da-57042b783b85.md)
