
# NavigationControl.Undo Event (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Occurs when the user undoes a change.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Undo**( **_Cancel_**, )

 _expression_A variable that represents a  **NavigationControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Cancel|Required| **Integer**||

### Return Value

nothing


## Remarks
<a name="sectionSection1"> </a>

The Undo event for controls occurs whenever the user returns a control to its original state by clicking the  **Undo Field/Record** button on the command bar, clicking the **Undo** button, pressing the ESC key, or calling the **Undo** method of the specified control. The control needs to have focus in all three cases. The event does not occur if the user clicks the **Undo Typing** button on the command bar.


## Example
<a name="sectionSection2"> </a>

The following example demonstrates the syntax for a subroutine that traps the Undo event for a form.


```
Private Sub Form_Undo(Cancel As Integer) 
 Dim intResponse As Integer 
 Dim strPrompt As String 
 
 strPrompt = "Cancel the undo operation?" 
 
 intResponse = MsgBox(strPrompt, vbYesNo) 
 
 If intResponse = vbYes Then 
 Cancel = True 
 Else 
 Cancel = False 
 End If 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [NavigationControl Object](ab08e35c-e5e4-444c-d169-1092d282ed15.md)
#### Other resources


 [NavigationControl Object Members](c972327e-9b46-f9fb-d69d-104d1d130ee4.md)
