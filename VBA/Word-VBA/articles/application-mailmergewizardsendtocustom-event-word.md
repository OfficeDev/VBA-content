---
title: Application.MailMergeWizardSendToCustom Event (Word)
keywords: vbawd10.chm4000022
f1_keywords:
- vbawd10.chm4000022
ms.prod: word
api_name:
- Word.Application.MailMergeWizardSendToCustom
ms.assetid: b5dcd912-f1b5-96d6-3221-d294211b6611
ms.date: 06/08/2017
---


# Application.MailMergeWizardSendToCustom Event (Word)

Occurs when the custom button is clicked during step six of the Mail Merge Wizard.


## Syntax

Private Sub  _expression_ _**MailMergeWizardSendToCustom**( **_ByVal Doc As Document_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object that has been declared with events in a class module.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The mail merge main document.|

## Remarks

Use the  **ShowSendToCustom** property to create a custom button on the sixth step of the Mail Merge Wizard.

For information about using events with the  **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).


## Example

This example executes a merge to a fax machine when a user clicks the custom destination button. This example assumes that the user has access to a custom destination button, fax numbers are included for each record in the data source, and an application variable called MailMergeApp has been declared and set equal to the Word Application object.


```vb
Private Sub MailMergeApp_MailMergeWizardSendToCustom(ByVal Doc As Document) 
 With Doc.MailMerge 
 .Destination = wdSendToFax 
 .Execute 
 End With 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

