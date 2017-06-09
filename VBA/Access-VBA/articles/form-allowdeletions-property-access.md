---
title: Form.AllowDeletions Property (Access)
keywords: vbaac10.chm13357,vbaac10.chm4260
f1_keywords:
- vbaac10.chm13357,vbaac10.chm4260
ms.prod: access
api_name:
- Access.Form.AllowDeletions
ms.assetid: abcbaa74-9a02-ab9c-613f-0cf6b9ce98b7
ms.date: 06/08/2017
---


# Form.AllowDeletions Property (Access)

You can use the  **AllowDeletions** property to specify whether a user can delete a record when using a form. Read/write **Boolean**.


## Syntax

 _expression_. **AllowDeletions**

 _expression_ A variable that represents a **Form** object.


## Remarks

You can set this property to No to allow users to view and edit existing records but not to delete them. When  **AllowDeletions** is set to Yes, records may be deleted so long as existing referential integrity rules aren't broken.

If you want to prevent changes to existing records (make a form read-only), set the  **[AllowAdditions](form-allowadditions-property-access.md)**, **AllowDeletions**, and **[AllowEdits](form-allowedits-property-access.md)** properties to No. You can also make records read-only by setting the **[RecordsetType](http://msdn.microsoft.com/library/a66d4043-08cc-ead1-f9ff-efc7d7ea21bf%28Office.15%29.aspx)** property to Snapshot.

When the  **AllowDeletions** property is set to No, the **Delete Record** command on the **Edit** menu isn't available.




 **Note**  When the Data Mode argument of the OpenForm action is set, Microsoft Access will override a number of form property settings. If the Data Mode argument of the OpenForm action is set to Edit, Microsoft Access will open the form with the following property settings:


-  **AllowEdits** — Yes
    
-  **AllowDeletions** — Yes
    
-  **AllowAdditions** — Yes
    
-  **DataEntry** — No
    
To prevent the OpenForm action from overriding any of these existing property settings, omit the Data Mode argument setting so that Microsoft Access will use the property settings defined by the form.


## Example

The following example examines the  **ControlType** property for all controls on a form. For each label and text box control, the procedure toggles the **SpecialEffect** property for those controls. When the label controls' **SpecialEffect** property is set to Shadowed and the text box controls' **SpecialEffect** property is set to Normal and the **AllowAdditions**, **AllowDeletions**, and **AllowEdits** properties are all set to **True**, the `intCanEdit` variable is toggled to allow editing of the underlying data.


```vb
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


#### Concepts


[Form Object](form-object-access.md)

