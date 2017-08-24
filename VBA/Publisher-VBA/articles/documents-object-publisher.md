---
title: Documents Object (Publisher)
keywords: vbapb10.chm8716287
f1_keywords:
- vbapb10.chm8716287
ms.prod: publisher
api_name:
- Publisher.Documents
ms.assetid: 855b1677-4072-1e17-c22c-6db08e0c7569
ms.date: 06/08/2017
---


# Documents Object (Publisher)

Represents all open publications. The  **Documents** collection contains all **Document** objects that are open in Microsoft Publisher.


## Example

Use the  **Documents** property to return the **Documents** collection. The following example lists all of the open publications.


```
Dim objDocument As Document 
Dim strMsg As String 
For Each objDocument In Documents 
 strMsg = strMsg &amp; objDocument.Name &amp; vbCrLf 
Next objDocument 
MsgBox Prompt:=strMsg, Title:="Current Documents Open", Buttons:=vbOKOnly
```

Use the  **Add** method to add a new document to the collection. A new and visible instance of Publisher is created when the **Add** method is called. The following example adds a new document to the **Documents** collection.




```
Dim objDocument As Document 
Set objDocument = Documents.Add 
With objDocument 
 .LayoutGuides.Columns = 4 
 .LayoutGuides.Rows = 9 
 .ActiveView.Zoom = pbZoomWholePage 
End With
```

Use the  **Item** (index) property, where index is the index number or document name as a **String**, to return a specific document object. The following example displays the name of the first open publication.




```
If Documents.Count >= 1 Then 
 MsgBox Documents.Item(1).Name 
End If 

```

The following example checks the name of each document in the  **Documents** collection. If the name of a document is "sales.doc", an object variable objSalesDoc is set to that document in the **Documents** collection.




```
Dim objDocument As Document 
Dim objSalesDoc As Document 
For Each objDocument In Documents 
 If objDocument.Name = "sales.pub" Then 
 Set objSalesDoc = objDocument 
 End If 
Next objDocument
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/1e3536c8-8fc0-8c95-3a4c-b16fe8a99098%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/41a2db19-5d56-be9b-a183-707d5e9e7e25%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/fe637a72-a96c-abfb-fa17-421848db5396%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/61cf3002-26d4-a678-abcb-940e7c385287%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/a0bca15f-39a0-f7f0-9b68-f6ba30414d50%28Office.15%29.aspx)|

