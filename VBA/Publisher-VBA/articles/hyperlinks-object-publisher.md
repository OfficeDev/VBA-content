---
title: Hyperlinks Object (Publisher)
keywords: vbapb10.chm6946815
f1_keywords:
- vbapb10.chm6946815
ms.prod: publisher
api_name:
- Publisher.Hyperlinks
ms.assetid: a82724b9-e792-b0e6-d1c3-25ce6021ad29
ms.date: 06/08/2017
---


# Hyperlinks Object (Publisher)

Represents the collection of  **[Hyperlink](hyperlink-object-publisher.md)** objects in a text range.


## Example

Use the  **[Hyperlinks](http://msdn.microsoft.com/library/0cf1f043-532c-3ffc-67cf-389adc5ac02f%28Office.15%29.aspx)** property to return the **Hyperlinks** collection. The following example deletes all text hyperlinks in the active publication that contain the word "Tailspin" in the address.


```
Sub DeleteMSHyperlinks() 
 Dim pgsPage As Page 
 Dim shpShape As Shape 
 Dim hprLink As Hyperlink 
 For Each pgsPage In ActiveDocument.Pages 
 For Each shpShape In pgsPage.Shapes 
 If shpShape.HasTextFrame = msoTrue Then 
 If shpShape.TextFrame.HasText = msoTrue Then 
 For Each hprLink In shpShape.TextFrame.TextRange.Hyperlinks 
 If InStr(hprLink.Address, "tailspin") <> 0 Then 
 hprLink.Delete 
 Exit For 
 End If 
 Next 
 Else 
 shpShape.Hyperlink.Delete 
 End If 
 End If 
 Next 
 Next 
End Sub
```

Use the  **[Add](http://msdn.microsoft.com/library/f5a8cc01-a571-623d-bfab-fe48e43a21b1%28Office.15%29.aspx)** method to create a hyperlink and add it to the **Hyperlinks** collection. The following example creates a new hyperlink to the specified Web site.




```
Sub AddHyperlink() 
 Selection.TextRange.Hyperlinks.Add Text:=Selection.TextRange, _ 
 Address:="http://www.tailspintoys.com/" 
End Sub
```

Use  **Hyperlinks** (index), where index is the index number, to return a single **Hyperlink** object in a publication, range, or selection. This example displays the address for the first hyperlink if the specified selection contains hyperlinks.




```
Sub DisplayHyperlinkAddress() 
 With Selection.TextRange.Hyperlinks 
 If .Count > 0 Then _ 
 MsgBox .Item(1).Address 
 End With 
End Sub
```

The  **[Count](http://msdn.microsoft.com/library/36747f3e-b365-11ca-9cbe-f6148f7da235%28Office.15%29.aspx)** property for this collection returns the number of hyperlinks in the specified shape or selection only.


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/f5a8cc01-a571-623d-bfab-fe48e43a21b1%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/c025e261-dc0e-9445-2c89-c9e79db6b3cd%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/36747f3e-b365-11ca-9cbe-f6148f7da235%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/8d288fc6-9ded-5732-b972-6fa366ef31c3%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/e3b25f19-6322-172a-3620-c3e728074655%28Office.15%29.aspx)|

