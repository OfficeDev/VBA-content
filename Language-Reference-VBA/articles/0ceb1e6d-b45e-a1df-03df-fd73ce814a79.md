
# StartUpPosition Property

 **Last modified:** July 28, 2015


Returns or sets a value specifying the position of a  **UserForm** when it first appears.
You can use one of four settings for  **StartUpPosition**:


|**Setting**|**Value**|**Description**|
|:-----|:-----|:-----|
| **Manual**|0|No initial setting specified.|
| **CenterOwner**|1|Center on the item to which the  **UserForm** belongs.|
| **CenterScreen**|2|Center on the whole screen.|
| **WindowsDefault**|3|Position in upper-left corner of screen.|
 **Remarks**
You can set the  **StartUpPosition** property programmatically or from the **Properties** window.

## Example

The following example uses the  **Load** statement and the **Show** method in UserForm1's Click event to load UserForm2 with the **StartUpPosition** property set to 3 (the Windows default position). The **Show** method then makes UserForm2 visible.


```
Private Sub UserForm_Click()
    Load UserForm2
    UserForm2. StartUpPosition = 3
    UserForm2.Show
End Sub
```

