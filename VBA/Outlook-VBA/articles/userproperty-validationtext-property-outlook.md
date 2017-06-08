---
title: UserProperty.ValidationText Property (Outlook)
keywords: vbaol11.chm221
f1_keywords:
- vbaol11.chm221
ms.prod: outlook
api_name:
- Outlook.UserProperty.ValidationText
ms.assetid: f2defd65-2c48-a24a-8cdc-a05b752cde53
ms.date: 06/08/2017
---


# UserProperty.ValidationText Property (Outlook)

Returns or sets a  **String** specifying the validation text for the specified user property. Read/write.


## Syntax

 _expression_ . **ValidationText**

 _expression_ A variable that represents a **UserProperty** object.


## Remarks

The validation text is the error message that a user receives when the  **[Value](userproperty-value-property-outlook.md)** does not meet the criteria specified in **[ValidationFormula](userproperty-validationformula-property-outlook.md)** .


## Example

The following Visual Basic for Applications (VBA) example demonstrates the use of  **ValidationText** and **ValidationFormula** properties.


```vb
Sub TestValidation() 
 
 Dim tki As Outlook.TaskItem 
 
 Dim uprs As Outlook.UserProperties 
 
 Dim upr As Outlook.UserProperty 
 
 
 
 Set tki = Application.CreateItem(olTaskItem) 
 
 tki.Subject = "Work hours" 
 
 ' TotalWork is stored in units of minutes 
 
 tki.TotalWork = 3000 
 
 Set uprs = tki.UserProperties 
 
 Set upr = uprs.Add("TotalWork", olFormula) 
 
 upr.Formula = "[Total Work]" 
 
 upr.ValidationFormula = ">= 2400" 
 
 upr.ValidationText = """The Work Hours (TotalWork) should be equal or greater than 5 days """ 
 
 tki.Save 
 
 tki.Display 
 
 MsgBox "The Work Hours are: " &; upr.Value / 60 
 
End Sub
```


## See also


#### Concepts


[UserProperty Object](userproperty-object-outlook.md)

