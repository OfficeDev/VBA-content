---
title: PostItem.CustomPropertyChange Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.PostItem.CustomPropertyChange
ms.assetid: e949151c-3824-1af2-6974-abd0d825e2c4
ms.date: 06/08/2017
---


# PostItem.CustomPropertyChange Event (Outlook)

Occurs when a custom property of an item (which is an instance of the parent object) is changed. 


## Syntax

 _expression_ . **CustomPropertyChange**( **_Name_** )

 _expression_ A variable that represents a **PostItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the custom property that was changed.|

## Remarks

The property name is passed to the procedure so that you can determine which custom property changed.


## Example

This Microsoft Visual Basic Scripting Edition (VBScript) example uses the  **CustomPropertyChange** event to enable a control when a **Boolean** field is set to **True** .

For this example, create two custom fields on the second page of a form. The first, a  **Boolean** field, is named "RespondBy". The second field is named "DateToRespond".




```vb
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


#### Concepts


[PostItem Object](postitem-object-outlook.md)

