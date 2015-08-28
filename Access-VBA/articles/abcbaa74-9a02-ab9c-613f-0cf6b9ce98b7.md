
# Form.AllowDeletions Property (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


You can use the  **AllowDeletions** property to specify whether a user can delete a record when using a form. Read/write **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **AllowDeletions**

 _expression_A variable that represents a  **Form** object.


## Remarks
<a name="sectionSection1"> </a>

You can set this property to No to allow users to view and edit existing records but not to delete them. When  **AllowDeletions** is set to Yes, records may be deleted so long as existing referential integrity rules aren't broken.

If you want to prevent changes to existing records (make a form read-only), set the  ** [AllowAdditions](8e440a96-7f9e-c009-5055-377c75999267.md)**,  **AllowDeletions**, and  ** [AllowEdits](3f667914-3dcc-7d4e-ca66-4338fc08e63a.md)**properties to No. You can also make records read-only by setting the  ** [RecordsetType](http://msdn.microsoft.com/library/a66d4043-08cc-ead1-f9ff-efc7d7ea21bf%28Office.15%29.aspx)**property to Snapshot.

When the  **AllowDeletions** property is set to No, the **Delete Record** command on the **Edit** menu isn't available.




 **Note**  When the Data Mode argument of the OpenForm action is set, Microsoft Access will override a number of form property settings. If the Data Mode argument of the OpenForm action is set to Edit, Microsoft Access will open the form with the following property settings:


-  **AllowEdits** â€” Yes
    
-  **AllowDeletions** â€” Yes
    
-  **AllowAdditions** â€” Yes
    
-  **DataEntry** â€” No
    
To prevent the OpenForm action from overriding any of these existing property settings, omit the Data Mode argument setting so that Microsoft Access will use the property settings defined by the form.


## Example
<a name="sectionSection2"> </a>

The following example examines the  **ControlType** property for all controls on a form. For each label and text box control, the procedure toggles the **SpecialEffect** property for those controls. When the label controls' **SpecialEffect** property is set to Shadowed and the text box controls' **SpecialEffect** property is set to Normal and the **AllowAdditions**,  **AllowDeletions**, and  **AllowEdits** properties are all set to **True**, the  `intCanEdit` variable is toggled to allow editing of the underlying data.


```
Sub ToggleControl(frm As Form) 
 Dim ctl As Control 
 Dim intI As Integer, intCanEdit As Integer 
 Const conTransparent = 0 
 Const conWhite = 16777215 
 For Each ctl in frm.Controls 
 With ctl 
 Select Case .ControlType 
 Case acLabel 
 If .SpecialEffect = acEffectShadow Then 
 .SpecialEffect = acEffectNormal 
 .BorderStyle = conTransparent 
 intCanEdit = True 
 Else 
 .SpecialEffect = acEffectShadow 
 intCanEdit = False 
 End If 
 Case acTextBox 
 If .SpecialEffect = acEffectNormal Then 
 .SpecialEffect = acEffectSunken 
 .BackColor = conWhite 
 Else 
 .SpecialEffect = acEffectNormal 
 .BackColor = frm.Detail.BackColor 
 End If 
 End Select 
 End With 
 Next ctl 
 If intCanEdit = IFalse Then 
 With frm 
 .AllowAdditions = False 
 .AllowDeletions = False 
 .AllowEdits = False 
 End With 
 Else 
 With frm 
 .AllowAdditions = True 
 .AllowDeletions = True 
 .AllowEdits = True 
 End With 
 End If 
End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Form Object](72ef9219-142b-b690-b696-3eba9a5d4522.md)
#### Other resources


 [Form Object Members](e1976b58-28ca-8f76-cdf3-6732cb06ce6c.md)
