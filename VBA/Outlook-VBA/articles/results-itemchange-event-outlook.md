---
title: Results.ItemChange Event (Outlook)
keywords: vbaol11.chm515
f1_keywords:
- vbaol11.chm515
ms.prod: outlook
api_name:
- Outlook.Results.ItemChange
ms.assetid: 14c96a47-00b8-6160-f1aa-386947ef50d1
ms.date: 06/08/2017
---


# Results.ItemChange Event (Outlook)

Occurs when an item in the specified collection is changed.


## Syntax

 _expression_ . **ItemChange**( **_Item_** )

 _expression_ A variable that represents a **Results** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The item that was changed.|

## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This example uses the  **[Start](appointmentitem-start-property-outlook.md)** property of the **[AppointmentItem](appointmentitem-object-outlook.md)** object to determine if the appointment starts after normal business hours. If it does, and if the **[Sensitivity](appointmentitem-sensitivity-property-outlook.md)** property of the **AppointmentItem** object is not already set to **olPrivate** , the example offers to mark the appointment as private.


```vb
Public WithEvents myOlItems As Outlook.Items 
 
 
 
Public Sub Initialize_handler() 
 
 Set myOlItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Items 
 
End Sub 
 
 
 
Private Sub myOlItems_ItemChange(ByVal Item As Object) 
 
Dim prompt As String 
 
 If VBA.Format(Item.Start, "h") >= "17" And Item.Sensitivity <> olPrivate Then 
 
 prompt = "Appointment occurs after hours. Mark it private?" 
 
 If MsgBox(prompt, vbYesNo + vbQuestion) = vbYes Then 
 
 Item.Sensitivity = olPrivate 
 
 Item.Display 
 
 End If 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Results Object](results-object-outlook.md)

