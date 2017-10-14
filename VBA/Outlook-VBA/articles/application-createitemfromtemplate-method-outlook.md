---
title: Application.CreateItemFromTemplate Method (Outlook)
keywords: vbaol11.chm715
f1_keywords:
- vbaol11.chm715
ms.prod: outlook
api_name:
- Outlook.Application.CreateItemFromTemplate
ms.assetid: 5e6c0ec4-779d-3743-afdb-606ad512ba95
ms.date: 06/08/2017
---


# Application.CreateItemFromTemplate Method (Outlook)

Creates a new Microsoft Outlook item from an Outlook template (.oft) and returns the new item.


## Syntax

 _expression_ . **CreateItemFromTemplate**( **_TemplatePath_** , **_InFolder_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TemplatePath_|Required| **String**|The path and file name of the Outlook template for the new item.|
| _InFolder_|Optional| **Variant**|The folder in which the item is to be created. If this argument is omitted, the default folder for the item type will be used.|

### Return Value

An  **Object** value that represents the new Outlook item.


## Remarks

New items will always open in compose mode, as opposed to read mode, regardless of the mode in which the items were saved to disk.


## Example

This Visual Basic for Applications (VBA) example uses  **CreateItemFromTemplate** to create a new item from an Outlook template and then displays it. The `CreateTemplate` macro shows you how to create the template that is used in the first example. To avoid errors, replace 'Dan Wilson' with a valid name in your address book.


```vb
Sub CreateFromTemplate() 
 Dim MyItem As Outlook.MailItem 
 
 Set MyItem = Application.CreateItemFromTemplate("C:\statusrep.oft") 
 MyItem.Display 
End Sub 
 
Sub CreateTemplate() 
 Dim MyItem As Outlook.MailItem 
 
 Set MyItem = Application.CreateItem(olMailItem) 
 MyItem.Subject = "Status Report" 
 MyItem.To = "Dan Wilson" 
 MyItem.Display 
 MyItem.SaveAs "C:\statusrep.oft", OlSaveAsType.olTemplate 
End Sub
```

The following Visual Basic for Applications (VBA) example shows how to use the optional  _InFolder_ parameter when calling the **CreateItemFromTemplate** method.




```vb
Sub CreateFromTemplate2() 
 Dim MyItem As Outlook.MailItem 
 
 Set MyItem = Application.CreateItemFromTemplate("C:\statusrep.oft", _ 
 Application.Session.GetDefaultFolder(olFolderDrafts)) 
 MyItem.Save 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-outlook.md)

