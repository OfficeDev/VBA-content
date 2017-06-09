---
title: Inspector.SetControlItemProperty Method (Outlook)
keywords: vbaol11.chm2980
f1_keywords:
- vbaol11.chm2980
ms.prod: outlook
api_name:
- Outlook.Inspector.SetControlItemProperty
ms.assetid: 90bb0dbf-c47e-9d75-182c-59c3e2384db2
ms.date: 06/08/2017
---


# Inspector.SetControlItemProperty Method (Outlook)

Binds a built-in property or custom property to a control in an inspector. 


## Syntax

 _expression_ . **SetControlItemProperty**( **_Control_** , **_PropertyName_** )

 _expression_ A variable that represents an **Inspector** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Control_|Required| **Object**|The control that will be bound to a property.|
| _PropertyName_|Required| **String**|The name of the property that will be bound to the control.|

## Remarks

You can use this method to bind an explicit built-in property or a custom property to a control. You must reference the property by its string name, for example,  **Subject** , and not by namespace, for example, http://schemas.microsoft.com/mapi/proptag/0x0037001E.

The  _PropertyName_ parameter is not case-sensitive. For example, **SetControlItemProperty** interprets an argument, _CustomerId_, to be the same as  _CustomerID_ and binds the specified control to the built-in **[ContactItem.CustomerID](contactitem-customerid-property-outlook.md)** property.

You can also use the following line of code  `myPage.Controls("bar").ItemProperty = "subject"` to bind the subject property to a control. However, note that untrusted code using this will trigger a security warning if the property is protected by the object model security guard such as **To** , and the client computer is running Microsoft Office Outlook 2007 or later but does not have an appropriately set up antivirus software. You can use the **SetControlItemProperty** method to avoid security warnings with trusted applications.


## Example

The following Visual Basic for Applications (VBA) code adds a custom page to an appointment item, adds a custom textbox control, and binds that control to  **Subject** property.


```vb
Sub Example() 
 Dim myIns As Outlook.Inspector 
 Dim myAppt As Outlook.AppointmentItem 
 Dim ctrl As Object 
 Dim ctrls As Object 
 Dim myPages As Outlook.Pages 
 Dim myPage As Object 
 
 Set myAppt = Application.CreateItem(olAppointmentItem) 
 Set myIns = myAppt.GetInspector 
 
 Set myPages = myIns.ModifiedFormPages 
 Set myPage = myPages.Add("New Page") 
 myIns.ShowFormPage ("New Page") 
 Set ctrls = myPage.Controls 
 Set ctrl = ctrls.Add("Forms.TextBox.1") 
 
 myIns.SetControlItemProperty ctrl, "Subject" 
 
 myAppt.Display 
End Sub
```


## See also


#### Concepts


[Inspector Object](inspector-object-outlook.md)

