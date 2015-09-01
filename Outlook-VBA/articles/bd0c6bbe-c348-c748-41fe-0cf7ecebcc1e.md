
# OlkDateControl Object (Outlook)

 **Last modified:** July 28, 2015

A control that supports the drop-down date picker used in inspectors for task and appointment items to select a date. 

## Remarks

Before you use this control for the first time in the forms designer, add the Microsoft Outlook Date Control to the control toolbox. You can only add this control to a form region in an Outlook form using the forms designer; you cannot add this control to a Visual Basic  **UserForm** object in the Visual Basic Editor.

The following is an example of the date control at runtime. This control supports Microsoft Windows themes.


![](../images/olDate_ZA10120280.gif)



This control can bind to any built-in or custom  **DateTime** field. However, the control does not support any date format setting for the field, nor does it support the select range behavior that is available in the appointment inspector.

If the  ** [Click](ec2483b8-0fe1-de86-dc01-9cafbde31e44.md)** event is implemented but the ** [DropButtonClick](425118d2-afa4-4582-1f89-857e5b7ae903.md)** event is not implemented, then clicking the drop button will fire only the **Click** event.

For more information about Outlook controls, see  [Controls in a Custom Form](fcba1b34-c526-5d01-8644-cb8852bd2348.md). For examples of add-ins in C# and Visual Basic .NET that use Outlook controls, see code sample downloads on MSDN. 


## See also


#### Concepts


 [Outlook Object Model Reference](73221b13-d8d8-99b8-3394-b95dbbfd5ddc.md)
#### Other resources


 [OlkDateControl Object Members](6bc09aee-2f4e-5042-a653-52c0c09068c5.md)
