---
title: Form.AllowEdits Property (Access)
keywords: vbaac10.chm13356
f1_keywords:
- vbaac10.chm13356
ms.prod: access
api_name:
- Access.Form.AllowEdits
ms.assetid: 3f667914-3dcc-7d4e-ca66-4338fc08e63a
ms.date: 06/08/2017
---


# Form.AllowEdits Property (Access)

You can use the  **AllowEdits** property to specify whether a user can edit saved records when using a form. Read/write **Boolean**.


## Syntax

 _expression_. **AllowEdits**

 _expression_ A variable that represents a **Form** object.


## Remarks

You can use the  **AllowEdits** property to prevent changes to existing data displayed by a form. If you want to prevent changes to data in a specific control, use the **Enabled** or **Locked** property.

If you want to prevent changes to existing records (make a form read-only), set the  **[AllowAdditions](form-allowadditions-property-access.md)**, **[AllowDeletions](form-allowdeletions-property-access.md)**, and **AllowEdits** properties to No. You can also make records read-only by setting the **[RecordsetType](http://msdn.microsoft.com/library/a66d4043-08cc-ead1-f9ff-efc7d7ea21bf%28Office.15%29.aspx)** property to Snapshot.

When the  **AllowEdits** property is set to No, the **Delete Record** and **Data Entry** menu commands aren't available for existing records. (They may still be available for new records if the **AllowAdditions** property is set to Yes.)

Changing a field value programmatically causes the current record to be editable, regardless of the  **AllowEdits** property setting. If you want to prevent the user from making changes to a record ( **AllowEdits** is No) that you need to edit programmatically, save the record after any programmatic changes; the **AllowEdits** property setting will be honored once again after any unsaved changes to the current record are saved.




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

