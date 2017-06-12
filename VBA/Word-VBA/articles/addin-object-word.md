---
title: AddIn Object (Word)
keywords: vbawd10.chm2430
f1_keywords:
- vbawd10.chm2430
ms.prod: word
api_name:
- Word.AddIn
ms.assetid: 5615a8a9-1fd6-04fa-1fee-ec16502bd84a
ms.date: 06/08/2017
---


# AddIn Object (Word)

Represents a single add-in, either installed or not installed. The  **AddIn** object is a member of the **[AddIns](addins-object-word.md)** collection. The **AddIns** collection contains all the add-ins available to Microsoft Word, regardless of whether they are currently loaded. The **AddIns** collection includes global templates or Word add-in libraries (WLLs) displayed in the **Templates and Add-ins** dialog box.


## Remarks

Use  **[AddIns](application-addins-property-word.md)** ( _index_ ), where _index_ is the add-in name or index number, to return a single **AddIn** object. You must exactly match the spelling (but not necessarily the capitalization) of the name, as it is shown in the **Templates and Add-Ins** dialog box. The following example loads the Letter.dot template as a global template.


```vb
AddIns("Letter.dot").Installed = True
```

The index number represents the position of the add-in in the list of add-ins in the  **Templates and Add-ins** dialog box. The following instruction displays the path of the first available add-in.




```vb
If Addins.Count >= 1 Then MsgBox Addins(1).Path
```

The following example creates a list of add-ins at the beginning of the active document. The list contains the name, path, and installed state of each available add-in.




```vb
With ActiveDocument.Range(Start:=0, End:=0) 
 .InsertAfter "Name" &; vbTab &; "Path" &; vbTab &; "Installed" 
 .InsertParagraphAfter 
 For Each oAddIn In AddIns 
 .InsertAfter oAddIn.Name &; vbTab &; oAddIn.Path &; vbTab _ 
 &; oAddIn.Installed 
 .InsertParagraphAfter 
 Next oAddIn 
 .ConvertToTable 
End With
```

Use the  **[Add](addins-add-method-word.md)** method to add an add-in to the list of available add-ins and (optionally) install it using the Install argument.




```
AddIns.Add FileName:="C:\Templates\Other\Letter.dot", Install:=True
```

To install an add-in shown in the list of available add-ins, use the  **[Installed](addin-installed-property-word.md)** property.




```vb
AddIns("Letter.dot").Installed = True
```


 **Note**  Use the  **[Compiled](addin-compiled-property-word.md)** property to determine whether an **AddIn** object is a template or a WLL.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


