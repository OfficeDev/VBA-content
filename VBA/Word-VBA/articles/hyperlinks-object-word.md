---
title: Hyperlinks Object (Word)
ms.prod: word
ms.assetid: 25801753-737f-9219-6a14-6531eb2ca699
ms.date: 06/08/2017
---


# Hyperlinks Object (Word)

Represents the collection of  **Hyperlink** objects in a document, range, or selection.


## Remarks

Use the  **Hyperlinks** property to return the **Hyperlinks** collection. The following example checks all the hyperlinks in document one for a link that contains the word "Microsoft" in the address. If a hyperlink is found, it is activated with the **Follow** method.


```
For Each hLink In Documents(1).Hyperlinks 
 If InStr(hLink.Address, "Microsoft") <> 0 Then 
 hLink.Follow 
 Exit For 
 End If 
Next hLink
```

Use the  **Add** method to create a hyperlink and add it to the **Hyperlinks** collection. The following example creates a new hyperlink to the MSN Web site.




```
ActiveDocument.Hyperlinks.Add Address:="http://www.msn.com/", _ 
 Anchor:=Selection.Range
```

Use  **Hyperlinks** (Index), where Index is the index number, to return a single **[Hyperlink](hyperlink-object-word.md)** object in a document, range, or selection. The following example activates the first hyperlink in the selection.




```
If Selection.HyperLinks.Count >= 1 Then 
 Selection.HyperLinks(1).Follow 
End If
```

The  **Count** property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the **Range** object.


## Methods



|**Name**|
|:-----|
|[Add](hyperlinks-add-method-word.md)|
|[Item](hyperlinks-item-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](hyperlinks-application-property-word.md)|
|[Count](hyperlinks-count-property-word.md)|
|[Creator](hyperlinks-creator-property-word.md)|
|[Parent](hyperlinks-parent-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
