---
title: Application.MailMergeRecipientListClose Event (Publisher)
keywords: vbapb10.chm268435488
f1_keywords:
- vbapb10.chm268435488
ms.prod: publisher
api_name:
- Publisher.Application.MailMergeRecipientListClose
ms.assetid: 4fb77771-9897-8623-f4e7-61f631f04922
ms.date: 06/08/2017
---


# Application.MailMergeRecipientListClose Event (Publisher)

Fires when the user closes the  **Mail Merge Recipients** dialog box. (From the **Mail Merge** or **E-mail Merge** task pane, click **Edit Recipient List**). Also fires when the user closes the  **Catalog Merge Product List** dialog box, which opens when the user clicks **Edit Product List** in the **Catalog Merge** task pane.


## Syntax

 _expression_. **MailMergeRecipientListClose**( **_Doc_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Doc|Required| **Document**|The current publication.|

## Remarks

For more information about using events with the  **Application** object, see [Using Events with the Application Object](using-events-with-the-application-object-publisher.md).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to handle the  **MailMergeRecipientListClose** event. It displays a message notifying the user that the string described above was displayed.


```vb
Private Sub pubApplication_MailMergeRecipientListClose(ByVal Doc As Document) 
 MsgBox "The Mail Merge Recipients dialog box has closed." 
End Sub
```

For this event to occur, you must place the following line of code in the  **General Declarations** section of your module.




```vb
Private WithEvents pubApplication As Application
```

Then run the following initialization procedure.




```vb
Public Sub Initialize_pubApplication() 
 Set pubApplication = Publisher.Application 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

