
# AllForms.Parent Property (Access)

 **Last modified:** July 28, 2015

Returns the parent object for the specified object. Read-only.

## Syntax

 _expression_. **Parent**

 _expression_A variable that represents an  **AllForms** object.


## Remarks

You can use the  **Parent** property to determine which form or report is currently the parent when you have a subform or subreport that has been inserted in multiple forms or reports.

For example, you might insert an OrderDetails subform into both a form and a report. The following example uses the  **Parent** property to refer to the OrderID field, which is present on the main form and report. You can enter this expression in a bound control on the subform.




```
=Parent!OrderID
```


## See also


#### Concepts


 [AllForms Collection](b90616b9-90fc-bb51-6bfa-b149dece0f1b.md)
#### Other resources


 [AllForms Object Members](a508646e-4478-fdfb-b1b5-177af651b73f.md)
