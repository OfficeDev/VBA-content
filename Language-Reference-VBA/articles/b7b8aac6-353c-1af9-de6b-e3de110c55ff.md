
# Page Object, CommandButton, MultiPage Controls, ControlTipText Property Example

 **Last modified:** July 28, 2015

The following example defines the  **ControlTipText** property for three **CommandButton** controls and two **Page** objects in a **MultiPage**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:



- A  **MultiPage** named MultiPage1.
    
- Three  **CommandButton** controls named CommandButton1 through CommandButton3.
    


 **Note**  For an individual  **Page** of a **MultiPage**,  **ControlTipText** becomes enabled when the **MultiPage** or a control on the current page of the **MultiPage** has the focus.




```
Private Sub UserForm_Initialize() 
 MultiPage1.Page1.ControlTipText = "Here in page 1" 
 MultiPage1.Page2.ControlTipText = "Now in page 2" 
 
 CommandButton1.ControlTipText = "And now here's" 
 CommandButton2.ControlTipText = "a tip from" 
 CommandButton3.ControlTipText = "your controls!" 
End Sub
```

