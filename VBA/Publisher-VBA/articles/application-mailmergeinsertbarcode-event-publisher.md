---
title: Application.MailMergeInsertBarcode Event (Publisher)
keywords: vbapb10.chm268435481
f1_keywords:
- vbapb10.chm268435481
ms.prod: publisher
api_name:
- Publisher.Application.MailMergeInsertBarcode
ms.assetid: 6b901953-eaff-0189-1d33-678e935a2f7e
ms.date: 06/08/2017
---


# Application.MailMergeInsertBarcode Event (Publisher)

Occurs when the user issues the command to insert postal barcodes into a mail-merge publication, either in the Microsoft Publisher user interface (UI), or programmatically.


## Syntax

 _expression_. **MailMergeInsertBarcode**( **_Doc_**,  **_OkToInsert_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Doc|Required| **Document**|The current publication.|
|OkToInsert|Required| **Boolean**|Output parameter.  **True** if it is okay to insert barcodes.|

## Remarks

You can use the  **[InsertBarcode](textrange-insertbarcode-method-publisher.md)** method to insert barcodes into a mail merge publication.

Third-party add-ins that validate mail-merge addresses can use the  **MailMergeInsertBarcode** event to listen for user actions requesting that barcodes be inserted. In this situation, when the add-in receives notification that the **MailMergeInsertBarcode** event fired, it checks the validity of the addresses in the mail-merge list, and if the addreses are valid, it attempts to generate barcodes. If this attempt is successful, the add-in should return **True** for the OkToInsert parameter. If the attempt fails, the add-in should return **False**.

Actual barcode data is provided to Publisher by the  **[MailMergeGenerateBarcode](application-mailmergegeneratebarcode-event-publisher.md)** event.

The  **MailMergeInsertBarcode** event is also triggered when a user clicks **Add a postal barcode** in the **Mail Merge** or **Catalog Merge** task pane, or **Add postal bar codes** in the **Publisher Tasks** task pane in the Publisher UI. Before a user can click either of these UI commands, you must first make them available by setting the **[InsertBarcodeVisible](application-insertbarcodevisible-property-publisher.md)** property to **True**. 

For more information about using events with the  **Application** object, see [Using Events with the Application Object](using-events-with-the-application-object-publisher.md).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to handle the  **MailMergeInsertBarcode** event. It displays a message asking whether to proceed with inserting barcodes.


```vb
Private Sub pubApplication_MailMergeInsertBarcode(ByVal Doc As Document, OkToInsert As Boolean) 
 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Proceed to insert barcodes?", vbYesNo) 
 
 If intResponse = vbYes Then OkToInsert = True 
 
End Sub
```

For this event to occur, you must place the following line of code in the  **General Declarations** section of your module.




```vb
Public WithEvents pubApplication As Application
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

