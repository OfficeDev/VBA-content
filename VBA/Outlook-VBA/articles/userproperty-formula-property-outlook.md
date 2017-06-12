---
title: UserProperty.Formula Property (Outlook)
keywords: vbaol11.chm217
f1_keywords:
- vbaol11.chm217
ms.prod: outlook
api_name:
- Outlook.UserProperty.Formula
ms.assetid: 91d2a104-8a93-a1e3-f31a-a0351153496d
ms.date: 06/08/2017
---


# UserProperty.Formula Property (Outlook)

Returns or sets a  **String** representing the formula for the user property. Read/write.


## Syntax

 _expression_ . **Formula**

 _expression_ A variable that represents a **UserProperty** object.


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **Formula** property.


```vb
Sub TestFormula() 
 
 Dim tki As Outlook.TaskItem 
 
 Dim uprs As Outlook.UserProperties 
 
 Dim upr As Outlook.UserProperty 
 
 
 
 Set tki = Application.CreateItem(olTaskItem) 
 
 tki.Subject = "Work hours - Test Formula" 
 
 ' TotalWork and ActualWork are in units of minutes 
 
 tki.TotalWork = 4 * 60 
 
 tki.ActualWork = 3 * 60 
 
 Set uprs = tki.UserProperties 
 
 Set upr = uprs.Add("Total&;ActualWork", olFormula) 
 
 upr.Formula = "[Total Work] + [Actual Work]" 
 
 tki.Save 
 
 tki.Display 
 
 MsgBox "The Work Hours are: " &; upr.Value / 60 
 
End Sub
```


## See also


#### Concepts


[UserProperty Object](userproperty-object-outlook.md)

