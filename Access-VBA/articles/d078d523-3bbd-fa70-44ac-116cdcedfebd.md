
# Report.Requery Method (Access)

 **Last modified:** July 28, 2015

The  **Requery** method updates the data underlying the specified report by requerying the source of data for the control.

## Syntax

 _expression_. **Requery**

 _expression_A variable that represents a  **Report** object.


## Remarks

You can use this method to ensure that a form or control displays the most recent data.

The  **Requery** method does one of the following:


- Reruns the query on which the report is based.
    
- Updates records displayed based on any changes to the  **Filter** property of the report.
    
If you omit the object specified by expression, the  **Requery** method requeries the underlying data source for the report that has the focus. If the control that has the focus has a record source or row source, it will be requeried; otherwise, the control's data will simply be refreshed.


 **Note**  


## See also


#### Concepts


 [Report Object](6f77c1b4-a9ce-7caa-204c-fe0755c6f9df.md)
#### Other resources


 [Report Object Members](73370a33-1ca0-da4d-9e36-88011bc2b93e.md)
