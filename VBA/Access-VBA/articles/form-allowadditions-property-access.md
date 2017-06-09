---
title: Form.AllowAdditions Property (Access)
keywords: vbaac10.chm13358
f1_keywords:
- vbaac10.chm13358
ms.prod: access
api_name:
- Access.Form.AllowAdditions
ms.assetid: 8e440a96-7f9e-c009-5055-377c75999267
ms.date: 06/08/2017
---


# Form.AllowAdditions Property (Access)

You can use the  **AllowAdditions** property to specify whether a user can add a record when using a form. Read/write **Boolean**.


## Syntax

 _expression_. **AllowAdditions**

 _expression_ A variable that represents a **Form** object.


## Remarks

Set the  **AllowAdditions** property to No to allow users to view or edit existing records but not add new records.

If you want to prevent changes to existing records (make a form read-only), set the  **AllowAdditions**, **[AllowDeletions](form-allowdeletions-property-access.md)**, and **[AllowEdits](form-allowedits-property-access.md)** properties to No. You can also make records read-only by setting the **[RecordsetType](http://msdn.microsoft.com/library/a66d4043-08cc-ead1-f9ff-efc7d7ea21bf%28Office.15%29.aspx)** property to Snapshot.

If you want to open a form for data entry only, set the form's  **[DataEntry](form-dataentry-property-access.md)** property to Yes.

When the  **AllowAdditions** property is set to No, the **Data Entry** command on the **Records** menu isn't available.


 **Note**  When the Data Mode argument of the OpenForm action is used, Microsoft Access will override a number of form property settings. If the Data Mode argument of the OpenForm action is set to Edit, Microsoft Access will open the form with the following property settings:


-  **AllowEdits** — Yes
    
-  **AllowDeletions** — Yes
    
-  **AllowAdditions** — Yes
    
-  **DataEntry** — No
    

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

