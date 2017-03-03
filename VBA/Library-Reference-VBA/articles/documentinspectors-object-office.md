---
title: DocumentInspectors Object (Office)
keywords: vbaof11.chm278000
f1_keywords:
- vbaof11.chm278000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.DocumentInspectors
ms.assetid: 8366d7cd-e016-bb99-d27f-749ca10352f1
---


# DocumentInspectors Object (Office)

Represents a collection of  **DocumentInspector** objects.


## Remarks

The  **DocumentInspectors** collection is part of the **Document** object in Microsoft Word, the **Workbook** object in Microsoft Excel, and the **Presentation** object in MicrosoftPowerPoint. A **DocumentInspectors** collection contains multiple **DocumentInspector** objects, one for some built-in options and each installed custom Document Inspector module. For more information, see the **DocumentInspector** help topic.


## Example

The following example calls the  **Fix** method of a Document Inspector module and displays the status of the action and the specific items that are removed.


```
Public Sub FixDocument() 
Dim docStatus As MsoDocInspectorStatus 
Dim results As String 
 ActiveDocument.DocumentInspectors(3).Fix docStatus, results 
 
 MsgBox docStatus 
 MsgBox("The following items were removed " &amp; results) 
 
End Sub 

```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/ea06ce71-5e18-1af3-2840-f1abeed4fbf1%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/78116c96-3d3e-2d91-a9a7-0826d16b2da6%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/cd22ea2b-5071-2ee1-abcd-32d7f06535e2%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/9f095ade-0e78-7158-b09e-ff068ebff20b%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/0d1f3b49-10ca-844c-6408-82d54842044e%28Office.15%29.aspx)|

## See also


#### Other resources


[DocumentInspectors Object Members](http://msdn.microsoft.com/library/1cf21432-076c-e5fe-496c-e20048a0e62e%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
