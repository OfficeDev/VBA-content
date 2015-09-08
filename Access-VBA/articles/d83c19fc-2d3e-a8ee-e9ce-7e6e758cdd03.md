
# BoundObjectFrame.Name Property (Access)

 **Last modified:** July 28, 2015

You can use the  **Name** property to specify or determine the string expression that identifies the name of an object. Read/write **String**.

## Syntax

 _expression_. **Name**

 _expression_A variable that represents a  **BoundObjectFrame** object.


## Remarks

A valid name must conform to the standard naming conventions for Microsoft Access. For Microsoft Access objects, the name may be up to 64 characters long. For controls, the name may be as long as 255 characters.

The default name for new objects is the object name plus a unique integer. For example, the first new form is Form1, the second new form is Form2, and so on. A form can't have the same name as another system object, such as the  ** [Screen](00743775-071b-9ccd-7687-f3b992e9346e.md)**object.

For an unbound control, the default name is the type of control plus a unique integer. For example, if the first control you add to a form is a text box, its  **Name** property setting is Text1

For a bound control, the default name is the name of the field in the underlying source of data. If you create a control by dragging a field from the field list, the field's  **FieldName**property setting is copied to the control's  **Name** property box.

You can't use "Form" or "Report" to name a control or section.

Controls on the same form, report, or data access page can't have the same name, but controls on different forms, reports or data access pages can have the same name. A control and a section on the same form can't share the same name.


## See also


#### Concepts


 [BoundObjectFrame Object](b3025672-60b8-e1d6-4769-1f724c9aa1ef.md)
#### Other resources


 [BoundObjectFrame Object Members](e2bbeb0c-1b13-5953-999a-4a0b93cb3ec7.md)
