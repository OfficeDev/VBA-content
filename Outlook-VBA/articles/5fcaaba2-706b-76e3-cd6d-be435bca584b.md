
# TaskRequestItem.CustomPropertyChange Event (Outlook)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Occurs when a custom property of an item (which is an instance of the parent object) is changed. 


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **CustomPropertyChange**( **_Name_**)

 _expression_A variable that represents a  **TaskRequestItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Name|Required| **String**|The name of the custom property that was changed.|

## Remarks
<a name="sectionSection1"> </a>

The property name is passed to the procedure so that you can determine which custom property changed.


## Example
<a name="sectionSection2"> </a>

This Microsoft Visual Basic Scripting Edition (VBScript) example uses the  **CustomPropertyChange**event to enable a control when a Boolean field is set to  **True**.

For this example, create two custom fields on the second page of a form. The first, a  **Boolean** field, is named "RespondBy". The second field is named "DateToRespond".




```
Sub Item_CustomPropertyChange(ByVal myPropName) 
 
 Select Case myPropName 
 
 Case "RespondBy" 
 
 Set myPages = Item.GetInspector.ModifiedFormPages 
 
 Set myCtrl = myPages("P.2").Controls("DateToRespond") 
 
 If Item.UserProperties("RespondBy").Value Then 
 
 myCtrl.Enabled = True 
 
 myCtrl.Backcolor = 65535 'Yellow 
 
 Else 
 
 myCtrl.Enabled = False 
 
 myCtrl.Backcolor = 0 'Black 
 
 End If 
 
 Case Else 
 
 End Select 
 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [TaskRequestItem Object](2908a28a-634c-e786-aa53-f3e32038b727.md)
#### Other resources


 [TaskRequestItem Object Members](d43114ee-be91-ff02-3424-525da2cf3a50.md)
