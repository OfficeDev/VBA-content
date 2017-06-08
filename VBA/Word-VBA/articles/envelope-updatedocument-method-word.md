---
title: Envelope.UpdateDocument Method (Word)
keywords: vbawd10.chm152567911
f1_keywords:
- vbawd10.chm152567911
ms.prod: word
api_name:
- Word.Envelope.UpdateDocument
ms.assetid: 6cca6549-58be-0b83-d52a-05fdccce0030
ms.date: 06/08/2017
---


# Envelope.UpdateDocument Method (Word)

Updates the envelope in the document with the current envelope settings.


## Syntax

 _expression_ . **UpdateDocument**

 _expression_ Required. A variable that represents an **[Envelope](envelope-object-word.md)** object.


## Remarks

If you use this property before an envelope has been added to the document, an error occurs.


## Example

This example formats the envelope in Report.doc to use a custom envelope size (4.5 inches by 7.5 inches).


```vb
Sub UpdateEnvelope() 
 
 On Error GoTo errhandler 
 
 With Documents("Report.doc").Envelope 
 .DefaultHeight = InchesToPoints(4.5) 
 .DefaultWidth = InchesToPoints(7.5) 
 .UpdateDocument 
 End With 
 
 Exit Sub 
 
errhandler: 
 
 If Err = 5852 Then _ 
 MsgBox "Report.doc doesn't include an envelope" 
 
End Sub
```

This example adds an envelope to the active document, using predefined addresses. The default envelope bar code and Facing Identification Mark (FIM-A) settings are set to  **True** , and the envelope in the active document is updated.




```vb
Dim strAddress As String 
Dim strReturn As String 
 
strAddress = "Darlene Rudd" &; vbCr &; "1234 E. Main St." _ 
 &; vbCr &; "Our Town, WA 98004" 
strReturn = "Patricia Reed" &; vbCr &; "N. 33rd St." _ 
 &; vbCr &; "Other Town, WA 98040" 
ActiveDocument.Envelope.Insert _ 
 Address:=strAddress, ReturnAddress:=strReturn 
With ActiveDocument.Envelope 
 .DefaultPrintBarCode = True 
 .DefaultPrintFIMA = True 
 .UpdateDocument 
End With
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

