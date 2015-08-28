
# Form.ApplyFilter Event (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Occurs when a filter is applied to a form.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ApplyFilter**( **_Cancel_**,  **_ApplyType_**)

 _expression_A variable that represents a  **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Cancel|Required| **Integer**|The setting determines if the  **ApplyFilter** event occurs. Setting theCancel argument to **True** cancels the **ApplyFilter** event and the filter is not applied to the form.|
|ApplyType|Required| **Integer**|Returns the type of filter that was applied.|

## Remarks
<a name="sectionSection1"> </a>

To run a macro or event procedure when this event occurs, set the  ** [OnApplyFilter](5e147a50-5516-f6d3-c1c9-e2c4522cb804.md)** property to the name of the macro or to [Event Procedure].

You can use the  **ApplyFilter** event to:


- Make sure the filter that is being applied is correct. For example, you may want to be sure that any filter applied to an Orders form includes criteria restricting the OrderDate field. To do this, check the form's  ** [Filter](5eb49f82-8519-981c-a663-9862736ac95f.md)** or ** [ServerFilter](18385de5-bc0d-9d2c-f97c-5b42e3689b45.md)** property value to make sure this criteria is included in the WHERE clause expression.
    
- Change the display of the form before the filter is applied. For example, when you apply a certain filter, you may want to disable or hide some fields that aren't appropriate for the records displayed by this filter.
    
- Undo or change actions you took when the Filter event occurred. For example, you can disable or hide some controls on the form when the user is creating the filter, because you don't want these controls to be included in the filter criteria. You can then enable or show these controls after the filter is applied. 
    
The actions in the  **ApplyFilter** macro or event procedure occur before the filter is applied or removed; or after the Advanced Filter/Sort, Filter By Form, or Server Filter By Form window is closed, but before the form is redisplayed. The criteria you've entered in the newly created filter are available to the **ApplyFilter** macro or event procedure as the setting of the **Filter** or **ServerFilter** property.


 **Note**  The  **ApplyFilter** event doesn't occur when the user does one of the following:


- Applies or removes a filter by using the  **ApplyFilter**,  **OpenReport**, or  **ShowAllRecords** actions in a macro, or their corresponding methods of the **DoCmd** object in Visual Basic.
    
- Uses the  **Close** action or the **Close** method of the **DoCmd** object to close the Advanced Filter/Sort, Filter By Form, or Server Filter By Form window
    
- Sets the  **Filter** or **ServerFilter** property or **FilterOn** or **ServerFilterByForm** property in a macro or Visual Basic (although you can set these properties in an **ApplyFilter** macro or event procedure).
    

## Example
<a name="sectionSection2"> </a>

The following example shows how to hide the AmountDue, Tax, and TotalDue controls on an Orders form when the applied filter restricts the records to only those orders that have been paid for.

To try this example, add the following event procedure to an Orders form that contains AmountDue, Tax, and TotalDue controls. Run a filter that lists only those orders that have been paid for.




```
Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer) 
 If Not IsNull(Me.Filter) And (InStr(Me.Filter, "Orders.Paid = -1")>0 _ 
 Or InStr(Me.Filter, "Orders.Paid = True")>0)Then 
 If ApplyType = acApplyFilter Then 
 Forms!Orders!AmountDue.Visible = False 
 Forms!Orders!Tax.Visible = False 
 Forms!Orders!TotalDue.Visible = False 
 ElseIf ApplyType = acShowAllRecords Then 
 Forms!Orders!AmountDue.Visible = True 
 Forms!Orders!Tax.Visible = True 
 Forms!Orders!TotalDue.Visible = True 
 End If 
 End If 
End Sub 

```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Form Object](72ef9219-142b-b690-b696-3eba9a5d4522.md)
#### Other resources


 [Form Object Members](e1976b58-28ca-8f76-cdf3-6732cb06ce6c.md)
