
# ComboBox.DisplayWhen Property (Access)

 **Last modified:** July 28, 2015

You can use the  **DisplayWhen** property to specify which of a form's controls you want displayed on screen and in print. Read/write **Byte**.

## Syntax

 _expression_. **DisplayWhen**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

The  **DisplayWhen** property applies only to the following form sections: detail, form header, and form footer. It also applies to all controls (except page breaks) on a form.

The  **DisplayWhen** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Always|0|(Default) The object appears in Form view and when printed.|
|Print Only|1|The object is hidden in Form view but appears when printed.|
|Screen Only|2|The object appears in Form view but not when printed.|
For controls, you can set the default for this property by using the default control style or the  **DefaultControl**property in Visual Basic.

In many cases, certain controls are useful only in Form view. To prevent Microsoft Access from printing these controls, you can set their  **DisplayWhen** property to Screen Only. For example, you might have a command button or instructions on a form that you don't want printed. Or you might have form header and form footer sections that you don't want displayed on screen but that you do want printed. In this case, you should set the **DisplayWhen** property to Print Only.

For reports, use the  **Format**and  **Retreat**events to specify an event procedure or macro that sets the  **Visible**property of controls you don't want printed. You can also cancel the Format or  **Print**event for a report section to prevent the section from being printed.


## See also


#### Concepts


 [ComboBox Object](1cf508d5-023e-eb38-3991-71e82b2a4e7e.md)
#### Other resources


 [ComboBox Object Members](d0d83ca3-3698-295e-5335-7d0816557d6b.md)
