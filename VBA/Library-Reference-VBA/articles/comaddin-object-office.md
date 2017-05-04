---
title: COMAddIn Object (Office)
keywords: vbaof11.chm219000
f1_keywords:
- vbaof11.chm219000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.COMAddIn
ms.assetid: dcaa9f0c-20fb-9f53-5f74-9ec0b1cefeea
---


# COMAddIn Object (Office)

Represents a COM add-in in the Microsoft Office host application. The  **COMAddIn** object is a member of the **COMAddIns** collection.


## Example

Use  **COMAddIns.Item(index)**, where _index_ is either an ordinal value that returns the COM add-in at that position in the **COMAddIns** collection, or a **String** value that represents the ProgID of the specified COM add-in. The following example displays a COM add-in's description text in a message box.


```
MsgBox Application.COMAddIns.Item("msodraa9.ShapeSelect").Description
```

Use the  **ProgID** property of the **COMAddin** object to return the programmatic identifier for a COM add-in, and use the **Guid** property to return the globally unique identifier (GUID) for the add-in. The following example displays the ProgID and GUID for COM add-in one in a message box.




```
MsgBox "My ProgID is " &amp; _ 
 Application.COMAddIns(1).ProgID &amp; _ 
 " and my GUID is " &amp; _ 
 Application.COMAddIns(1).Guid
```

Use the  **Connect** property to set or return the state of the connection to a specified COM add-in. The following example displays a message box that indicates whether COM add-in one is registered and currently connected.




```
If Application.COMAddIns(1).Connect Then 
 MsgBox "The add-in is connected." 
Else 
MsgBox "The add-in is not connected." 
End If
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/7566c80d-a63b-2ea0-7a53-21c532039172%28Office.15%29.aspx)|
|[Connect](http://msdn.microsoft.com/library/b1392380-c19f-ab3e-c9dc-c62438b16500%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/512057ba-021f-cb14-1123-e6d4061cca3e%28Office.15%29.aspx)|
|[Description](http://msdn.microsoft.com/library/f194ae48-0762-732f-7c9a-f19a92e94d9b%28Office.15%29.aspx)|
|[Guid](http://msdn.microsoft.com/library/1e3218d9-dce7-21e2-55a7-4435ca58bb35%28Office.15%29.aspx)|
|[Object](http://msdn.microsoft.com/library/20dd8eca-6f8e-7445-ec0c-a29b29409c58%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/dddd8d3a-f5d7-7a30-8301-f9dd0775f0a8%28Office.15%29.aspx)|
|[ProgId](http://msdn.microsoft.com/library/eb917d53-512e-35dd-ff70-ac7b976e6500%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[COMAddIn Object Members](http://msdn.microsoft.com/library/698d4d8e-6071-acd3-a39b-ab01fd878452%28Office.15%29.aspx)
