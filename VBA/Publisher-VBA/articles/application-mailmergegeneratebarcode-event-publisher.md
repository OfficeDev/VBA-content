---
title: Application.MailMergeGenerateBarcode Event (Publisher)
keywords: vbapb10.chm268435489
f1_keywords:
- vbapb10.chm268435489
ms.prod: publisher
api_name:
- Publisher.Application.MailMergeGenerateBarcode
ms.assetid: 5da4ec65-32b6-ea05-09ad-d2224eafee30
ms.date: 06/08/2017
---


# Application.MailMergeGenerateBarcode Event (Publisher)

Occurs when Microsoft Publisher requires data to generate barcodes in a mail-merge publication, in particular when the mail-merge recipient list changes.


## Syntax

 _expression_. **MailMergeGenerateBarcode**( **_Doc_**,  **_bstrString_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Doc|Required| **Document**|The current publication.|
|bstrString|Required| **String**|Output parameter. A string representation of the barcode.|

## Remarks

Third-party add-ins that validate mail-merge addresses can use the  **MailMergeGenerateBarcode** event to listen for user actions requesting that barcodes be generated. In this situation, when the add-in receives notification that the **MailMergeGenerateBarcode** event fired, and if the active document is connected to a data source, the add-in can use the ** [MailMergeDataSource.ActiveRecord](mailmergedatasource-activerecord-property-publisher.md)** property to determine the record for which to generate the barcode. If the active document is not connected to a data source, the add-in uses the address text directly.

If the add-in can use the address text directly, it returns a string representation of the barcode for the bstrString output parameter. If the add-in cannot use the address text directly, it returns an empty string.

To permit triggering of the  **MailMergeGenerateBarcode** event, you must handle the **[MailMergeInsertBarcode](application-mailmergeinsertbarcode-event-publisher.md)** event in your code, and the add-in must set the OkToInsert parameter passed to that event to **True**. 

For more information about using events with the  **Application** object, see [Using Events with the Application Object](using-events-with-the-application-object-publisher.md).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to handle the  **MailMergeGenerateBarcode** event. It returns the string that represents the barcode for active record. Note that the variable _indexNumberOfBarcodeColumn_ represents the index number of the column in the data source that lists barcodes. This code assumes that the current publication is connected to a data source.


```vb
Private Sub pubApplication_MailMergeGenerateBarcode(ByVal Doc As Document, bstrString As String) 
 bstrString = pubApplication.ActiveDocument.MailMerge.DataSource.DataFields.Item(indexNumberOfBarcodeColumn).Value 
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

