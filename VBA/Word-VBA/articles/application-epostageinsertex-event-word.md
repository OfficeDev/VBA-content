---
title: Application.EPostageInsertEx Event (Word)
keywords: vbawd10.chm4000028
f1_keywords:
- vbawd10.chm4000028
ms.prod: word
api_name:
- Word.Application.EPostageInsertEx
ms.assetid: 494225b9-f55f-37d2-8ff0-086f8d917b05
ms.date: 06/08/2017
---


# Application.EPostageInsertEx Event (Word)

Occurs when a user inserts electronic postage into a document.


## Syntax

 _expression_ . **EPostageInsertEx**( **_Doc_** , **_cpDeliveryAddrStart_** , **_cpDeliveryAddrEnd_** , **_cpReturnAddrStart_** , **_cpReturnAddrEnd_** , **_xaWidth_** , **_yaHeight_** , **_bstrPrinterName_** , **_bstrPaperFeed_** , **_fPrint_** , **_fCancel_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object that has been declared with events in a class module. For information about using events with the **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The document to which electronic postage is being added.|
| _cpDeliveryAddrStart_|Required| **Long**|The starting position in the document for the delivery address. Positioning corresponds to the value of the  **Start** property for a **Range** object.|
| _cpDeliveryAddrEnd_|Required| **Long**|The ending position in the document for the delivery address. Positioning corresponds to the value of the  **End** property for a **Range** object.|
| _cpReturnAddrStart_|Required| **Long**|The starting position in the document for the return address. Positioning corresponds to the value of the  **Start** property for a **Range** object.|
| _cpReturnAddrEnd_|Required| **Long**|The ending position in the document for the return address. Positioning corresponds to the value of the  **End** property for a **Range** object.|
| _xaWidth_|Required| **Long**|The width of the envelope in 1/1440-inch units.|
| _yaHeight_|Required| **Long**|The height of the envelope in 1/1440-inch units.|
| _bstrPrinterName_|Required| **String**|The name of the printer as specified on the  **Printing Options** tab of the **Envelope Options** dialog box.|
| _bstrPaperFeed_|Required| **String**|The feed method as specified on the  **Printing Options** tab of the **Envelope Options** dialog box.|
| _fPrint_|Required| **Boolean**| **True** if the user has specified to print the envelope. **False** if the user has specified to insert the envelope into the document.|
| _fCancel_|Required| **Boolean**| **True** cancels inserting the postage.|

## Example

The following example displays a message to the user. If the user cancels the message, then the action specified by the user is canceled.


```vb
Private Sub App_EPostageInsertEx(ByVal Doc As Document, ByVal cpDeliveryAddrStart As Long, _ 
 ByVal cpDeliveryAddrEnd As Long, ByVal cpReturnAddrStart As Long, _ 
 ByVal cpReturnAddrEnd As Long, ByVal xaWidth As Long, ByVal yaHeight As Long, _ 
 ByVal bstrPrinterName As String, ByVal bstrPaperFeed As String, _ 
 ByVal fPrint As Boolean, fCancel As Boolean) 
 
 Dim intResponse As Integer 
 
 If fPrint = True Then 
 intResponse = MsgBox("Make sure the printer is ready to print an envelope." &; vbCrLf &; _ 
 "When the printer is ready, click OK.", vbOKCancel) 
 
 If intResponse = vbCancel Then 
 fCancel = True 
 End If 
 End If 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

