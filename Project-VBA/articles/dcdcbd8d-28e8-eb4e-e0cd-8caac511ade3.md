
# ReportTable.GetCellText Method (Project)
Returns the text value of the specified cell in a  **ReportTable** object.

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **GetCellText**(Row,Col)

 _expression_A variable that represents a  **ReportTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|Row|Required| **Long**|The row number in the table.|
|Col|Required| **Long**|The column number in the table.|
|Row|Required|INT||
|Col|Required|INT||

### Return value

 **String**

The text value of the specified table cell.


## Remarks
<a name="sectionSection1"> </a>

The returned string ends with a newline character ( `chr(10)`, which is equivalent to the  **vbCrLf** character).


## Example
<a name="sectionSection2"> </a>

The  **GetTableText** example finds all of the tables on the active report, gets the value of each cell in a table, removes the last character of each value (the newline character), and then prints the table cell values to the Immediate window in the VBE. To use the **GetTableText** macro, create a project with values such as the example that is specified in the [Chart Object](810d4ec1-69d2-c432-b9da-57042b783b85.md) topic, and then do the following steps (see Figure 1):


1. Manually create a report. For example, on the  **PROJECT** tab of the ribbon, in the **Reports** drop-down list, choose **More Reports**. In the  **Reports** dialog box, choose **New** in the left pane, choose **Blank** in the right pane, and then choose **Select**. In the  **Report Name** dialog box, typeReport 1.
    
2. Add two tables to the report. Under  **REPORT TOOLS** on the **DESIGN** tab of the ribbon, use the **Table** command in the **Insert** group.
    
3. Keep the default values in the first table, which includes the  **Name**,  **Start**,  **Finish**, and  **% Complete** fields of the project summary task. Select the first table to display the **Field List** task pane, and then select **Actual Cost** and **Remaining Cost**.
    
4. Select the second table. In the  **Field List** task pane, change the **Filter** to **All Tasks**, and then select  **Actual Cost** and **Remaining Cost**. In the table, select and delete the  **Start** column and the **Finish** column.
    
5. Add two text boxes to the report, by using the  **Text Box** control in the **Insert** group on the ribbon. For example, edit the first text box to showProject summary task, and edit the second text box to show Task information.
    

**Figure 1. The sample report contains two tables and three text boxes**

![Report with two tables and three text boxes](../images/pj15_VBA_ReportTable_GetCellText.gif)




```
Sub GetTableText()
    Dim theReport As Report
    Dim shp As shape
    Dim theReportTable As ReportTable
    Dim reportName As String
    Dim row As Integer, col As Integer, i As Integer
    Dim output As String
    
    reportName = "Report 1"
    
    For i = 1 To ActiveProject.Reports(reportName).Shapes.Count
        Set shp = ActiveProject.Reports(reportName).Shapes(i)
        Debug.Print shp.Name &amp; "; ID = " &amp; shp.ID
    Next i
    
    For Each shp In ActiveProject.Reports(reportName).Shapes
        If shp.HasTable Then
            Debug.Print vbCrLf &amp; "Table name: " &amp; shp.Name
            
            For row = 1 To shp.Table.RowsCount
                output = vbTab
                
                For col = 1 To shp.Table.ColumnsCount
                    output = output &amp; shp.Table.GetCellText(row, col)
                    output = left(output, Len(output) - 1) &amp; vbTab
                Next col
                
                Debug.Print output
            Next row
        End If
    Next shp
End Sub
```

When you run the  **GetTableText** macro, the Immediate window in the VBE shows the following text. The top five lines show how shape objects are named by default and how ID values are created.




```
TextBox 1; ID = 2
Table 2; ID = 3
Table 3; ID = 4
TextBox 4; ID = 5
TextBox 5; ID = 6

Table name: Table 2
    Name    Start   Finish  % Complete  Actual Cost Remaining Cost  
    TestShapes  Mon 5/14/12 Tue 5/31/12 58% $1,595.00   $2,125.00   

Table name: Table 3
    Name    % Complete  Actual Cost Remaining Cost  
    T1  100%    $0.00   $0.00   
    T2  71% $1,280.00   $640.00 
    T3  44% $315.00 $765.00 
    T4  0%  $0.00   $720.00
```


## See also
<a name="sectionSection2"> </a>


#### Other resources


 [ReportTable Object](db9846c7-fd53-ae5a-7a43-35dfc60f4fe4.md)
 [ID Property](8b619251-1914-cbf0-6b50-e978f8ffe125.md)
