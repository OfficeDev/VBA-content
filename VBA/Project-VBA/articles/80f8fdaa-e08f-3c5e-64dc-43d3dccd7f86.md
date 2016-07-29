
# Application.GanttBarLinks Method (Project)

Shows or hides task links on the Gantt Chart.


## Syntax

 _expression_. **GanttBarLinks**( ** _Display_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Display_|Optional|**Long**|Where links will be drawn from the ends of predecessor links. Can be one of the  **[PjGanttBarLink](aa55b82c-f639-ad1d-b156-861f006267f4.md)** constants. The default value is **PjNoGanttBarLinks**.|

### Return Value

 **Boolean**


## Example

The following example first clears the links and then displays them from the end of one task bar to the top of the next task bar.


```vb
Sub GanttBar_Links() 
'First clear links, than links from end to top of the next bar 
 'Activate Gantt Chart view 
 ViewApply Name:="&;Gantt Chart" 
 GanttBarLinks pjNoGanttBarLinks 
 GanttBarLinks pjToTop 
End Sub
```

