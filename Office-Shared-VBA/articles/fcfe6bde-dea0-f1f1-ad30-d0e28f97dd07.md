
# CommandBarComboBox Object (Office)

 **Last modified:** July 28, 2015

 **In this article**
 [](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Represents a combo box control on a command bar.


## 
<a name="sectionSection0"> </a>


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Remarks
<a name="sectionSection1"> </a>

Use  **Controls(index)**, where  _index_ is the index number of the control, to return a **CommandBarComboBox** object. Note that the **Type** property of the control must be **msoControlEdit**,  **msoControlDropdown**,  **msoControlComboBox**,  **msoControlButtonDropdown**,  **msoControlSplitDropdown**,  **msoControlOCXDropdown**,  **msoControlGraphicCombo**, or  **msoControlGraphicDropdown**.


## Example
<a name="sectionSection2"> </a>

The following example adds two items to the second control on the command bar named  **Custom**, and then it adjusts the size of the control.


```
Set combo = CommandBars("Custom").Controls(2) 
With combo 
    .AddItem "First Item", 1 
    .AddItem "Second Item", 2 
    .DropDownLines = 3 
    .DropDownWidth = 75 
    .ListIndex = 0 
End With
```

You can also use the  **FindControl** method to return a **CommandBarComboBox** object. The following example searches all command bars for a visible **CommandBarComboBox** object whose tag is "sheet assignments."




```
Set myControl = CommandBars.FindControl _ 
(Type:=msoControlComboBox, Tag:="sheet assignments", Visible:=True)
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Object Model Reference](499c789a-aba2-0fad-649a-0ea964cd3b5e.md)
#### Other resources


 [CommandBarComboBox Object Members](223c51c0-4564-d14a-a8bf-d315a6a50b32.md)
