---
title: TextRange.InsertBarcode Method (Publisher)
keywords: vbapb10.chm5308502
f1_keywords:
- vbapb10.chm5308502
ms.prod: publisher
api_name:
- Publisher.TextRange.InsertBarcode
ms.assetid: ad613ca7-f056-55b0-1a96-51167555ce6f
ms.date: 06/08/2017
---


# TextRange.InsertBarcode Method (Publisher)

Inserts a bar code field at the end of the text range represented by the parent  **TextRange** object.


## Syntax

 _expression_. **InsertBarcode**

 _expression_A variable that represents a  **TextRange** object.


### Return Value

TextRange


## Remarks

Ideally, you should create an add-in to Microsoft Publisher to handle the  **[MailMergeGenerateBarcode](application-mailmergegeneratebarcode-event-publisher.md)** and **[MailMergeInsertBarcode](application-mailmergeinsertbarcode-event-publisher.md)** events. If your add-in or code does not contain handlers for these events, the **InsertBarcode** method returns an error.

The example that follows shows how to handle these events by using Microsoft Visual Basic for Applications (VBA) code in the Visual Basic Editor.

If you want to enable insertion of bar codes into the publication from the user interface, your add-in or VBA code should also set the  **[InsertBarcodeVisible](application-insertbarcodevisible-property-publisher.md)** property value to **True**.


## Example

The following example shows how to use the  **InsertBarcode** method to insert a bar-code field into a text box in a publication. Insert this code into your VBA project, and run the **AttachToEvents** procedure before running the **InsertBarcode_Example** procedure.

Before running the code in this example, use the  ** [MailMerge.OpenDataSource](mailmerge-opendatasource-method-publisher.md)** method to connect to a data source. The data source must contain a bar-code column that lists bar codes for all mail-merge recipients. Replace _barcodeColumnIndex_ in the **MailMergeGenerateBarcode** event handler in the code with the index number of the data-source column that contains bar-code information.

Run the following code from the  **Visual Basic Editor** window, and not from the **Macros** dialog box. (On the **Tools** menu, point to **Macro**, and then click Macros.)




```vb
Public WithEvents pubApplication As Publisher.Application 
 
Private Sub pubApplication_MailMergeGenerateBarcode(ByVal Doc As Document, bstrString As String) 
 
    bstrString = pubApplication.ActiveDocument.MailMerge.DataSource.DataFields.Item(barcodeColumnIndex).Value 
         
End Sub 
 
Private Sub pubApplication_MailMergeInsertBarcode(ByVal Doc As Document, OkToInsert As Boolean) 
 
    OkToInsert = True 
     
End Sub 
 
Public Sub InsertBarcode_Example() 
 
    Dim pubTextRange As Publisher.TextRange 
    Dim pubShape As Publisher.Shape 
     
    Set pubShape = ThisDocument.Pages(1).Shapes.AddTextbox(pbTextOrientationHorizontal, 100, 100, 500, 500) 
    Set pubTextRange = pubShape.TextFrame.TextRange 
     
    pubTextRange.InsertBarcode 
     
End Sub 
 
Public Sub AttachToEvents() 
 
    Set pubApplication = Application 
 
End Sub
```


