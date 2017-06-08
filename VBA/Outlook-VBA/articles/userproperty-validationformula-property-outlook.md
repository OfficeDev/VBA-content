---
title: UserProperty.ValidationFormula Property (Outlook)
keywords: vbaol11.chm220
f1_keywords:
- vbaol11.chm220
ms.prod: outlook
api_name:
- Outlook.UserProperty.ValidationFormula
ms.assetid: 1420a7d9-2d10-ea1a-a893-e573f93919ad
ms.date: 06/08/2017
---


# UserProperty.ValidationFormula Property (Outlook)

Returns or sets a  **String** indicating the validation formula for the user property. Read/write.


## Syntax

 _expression_ . **ValidationFormula**

 _expression_ A variable that represents a **UserProperty** object.


## Remarks

The validation formula is used by Outlook to validate the  **[Value](userproperty-value-property-outlook.md)** property when an item is saved.


## Example

The following Visual Basic for Applications (VBA) example demonstrates the use of  **ValidationText** and **ValidationFormula** properties.


```vb
Sub TestValidation() 
 
 Dim tki As Outlook.TaskItem 
 
 Dim uprs As Outlook.UserProperties 
 
 Dim upr As Outlook.UserProperty 
 
 
 
 Set tki = Application.CreateItem(olTaskItem) 
 
 tki.Subject = "Work hours" 
 
 tki.TotalWork = 3000 
 
 Set uprs = tki.UserProperties 
 
 Set upr = uprs.Add("TotalWork", olFormula) 
 
 upr.Formula = "[Total Work]" 
 
 upr.ValidationFormula = ">= 2400" 
 
 upr.ValidationText = """The WorkHours (Total Work) should be equal or greater than 5 days """ 
 
 tki.Save 
 
 tki.Display 
 
 MsgBox "The Work Hours are: " &; upr.Value 
 
End Sub
```


## See also


#### Concepts


[UserProperty Object](userproperty-object-outlook.md)

