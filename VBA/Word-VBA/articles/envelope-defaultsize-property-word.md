---
title: Envelope.DefaultSize Property (Word)
keywords: vbawd10.chm152567808
f1_keywords:
- vbawd10.chm152567808
ms.prod: word
api_name:
- Word.Envelope.DefaultSize
ms.assetid: 2365a10b-229c-141b-49ab-7d6a0e2247b2
ms.date: 06/08/2017
---


# Envelope.DefaultSize Property (Word)

Returns or sets the default envelope size. Read/write  **String** .


## Syntax

 _expression_ . **DefaultSize**

 _expression_ A variable that represents a **[Envelope](envelope-object-word.md)** object.


## Remarks

The string that is returned corresponds to the right side of the string that appears in the  **Envelope Size** box in the **Envelope Options** dialog box. If you set either the **[DefaultHeight](envelope-defaultheight-property-word.md)** or **[DefaultWidth](envelope-defaultwidth-property-word.md)** property, the envelope size is automatically changed to **Custom Size** in the **Envelope Options** dialog box ( **Tools** menu) and this property returns "Custom size."


## Example

This example sets the default envelope size to C4 (229 x 324 mm).


```vb
ActiveDocument.Envelope.DefaultSize = "C4"
```

This example asks the user whether or not they want to change the default envelope size to Size 10. If the answer is yes, the default size is changed accordingly. The UpdateDocument method changes the envelope size for the active document. If an envelope has not been added to the active document, a message box is displayed.




```vb
Sub exDefaultSize() 
 
 Dim intResponse As Integer 
 
 On Error GoTo errhandler 
 intResponse = MsgBox("Do you want to set the " _ 
 &; "default envelope to Size 10?", 4) 
 If intResponse = vbYes Then 
 With ActiveDocument.Envelope 
 .DefaultSize = "Size 10" 
 .UpdateDocument 
 End With 
 End If 
 
 Exit Sub 
 
errhandler: 
 If Err = 5852 Then _ 
 MsgBox "An envelope isn't part of this document" 
End Sub
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

