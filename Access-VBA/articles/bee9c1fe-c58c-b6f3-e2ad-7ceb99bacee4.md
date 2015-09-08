
# SubForm.SourceObject Property (Access)

 **Last modified:** July 28, 2015

You can use the  **SourceObject** property to identify the form or report that is the source of the subform or subreport on a form or report. Read/write **String**.

## Syntax

 _expression_. **SourceObject**

 _expression_A variable that represents a  **SubForm** object.


## Remarks

Enter the name of the form or report that is the source of the subform or subreport in the control's property sheet. If you add a subform or subreport to the form or report by dragging it from the Database window, the  **SourceObject** property is set automatically in the property sheet.

In Visual Basic, you set this property by using a string expression that is a name of a form or report.


 **Note**  You can't set or change the  **SourceObject** property in the **Open**or  **Format**events of a report.

If you delete the  **SourceObject** property setting in the property sheet for a subform or subreport, the control remains on the form but is no longer bound to the source form or report.


## See also


#### Concepts


 [SubForm Object](60f961fa-dcf4-e1d1-8c50-9e88963f9dec.md)
#### Other resources


 [SubForm Object Members](328e74d8-0418-968f-faca-3e1b34139f48.md)
