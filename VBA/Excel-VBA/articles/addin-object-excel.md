---
title: AddIn Object (Excel)
keywords: vbaxl10.chm184072
f1_keywords:
- vbaxl10.chm184072
ms.prod: excel
api_name:
- Excel.AddIn
ms.assetid: ad26800d-5342-fb4c-01f3-05b7eceb7ffd
ms.date: 06/08/2017
---


# AddIn Object (Excel)

Represents a single add-in, either installed or not installed. 


## Remarks

The  **AddIn** object is a member of the **[AddIns](addins-object-excel.md)** collection. The **AddIns** collection contains a list of all the add-ins available to Microsoft Excel, regardless of whether they're installed. This list corresponds to the list of add-ins displayed in the **Add-Ins** dialog box.


## Example

Use  **AddIns** ( _index_ ), where _index_ is the add-in title or index number, to return a single **AddIn** object. The following example installs the Analysis Toolpak add-in.


```
AddIns("analysis toolpak").Installed = True
```

Don't confuse the add-in title, which appears in the  **Add-Ins** dialog box, with the add-in name, which is the file name of the add-in. You must spell the add-in title exactly as it's spelled in the **Add-Ins** dialog box, but the capitalization doesn't have to match.

The index number represents the position of the add-in in the  **Add-ins available** box in the **Add-Ins** dialog box. The following example creates a list that contains specified properties of the available add-ins.




```
With Worksheets("sheet1") 
 .Rows(1).Font.Bold = True 
 .Range("a1:d1").Value = _ 
 Array("Name", "Full Name", "Title", "Installed") 
 For i = 1 To AddIns.Count 
 .Cells(i + 1, 1) = AddIns(i).Name 
 .Cells(i + 1, 2) = AddIns(i).FullName 
 .Cells(i + 1, 3) = AddIns(i).Title 
 .Cells(i + 1, 4) = AddIns(i).Installed 
 Next 
 .Range("a1").CurrentRegion.Columns.AutoFit 
End With
```

The  **[Add](addins-add-method-excel.md)** method adds an add-in to the list of available add-ins but doesn't install the add-in. Set the **[Installed](addin-installed-property-excel.md)** property of the add-in to **True** to install the add-in. To install an add-in that doesn't appear in the list of available add-ins, you must first use the **Add** method and then set the **Installed** property. This can be done in a single step, as shown in the following example (note that you use the name of the add-in, not its title, with the **Add** method).




```
AddIns.Add("generic.xll").Installed = True
```

Use  **Workbooks** ( _index_ ) where _index_ is the add-in filename (not title) to return a reference to the workbook corresponding to a loaded add-in. You must use the file name because loaded add-ins don't normally appear in the **Workbooks** collection. This example sets the _wb_ variable to the workbook for Myaddin.xla.




```
Set wb = Workbooks("myaddin.xla")
```

The following example sets the  _wb_ variable to the workbook for the Analysis Toolpak add-in.




```
Set wb = Workbooks(AddIns("analysis toolpak").Name)
```

If the  **Installed** property returns **True**, but calls to functions in the add-in still fail, the add-in may not actually be loaded. This is because the **Addin** object represents the existence and installed state of the add-in but doesn't represent the actual contents of the add-in workbook.To guarantee that an installed add-in is loaded, you should open the add-in workbook. The following example opens the workbook for the add-in named "My Addin" if the add-in isn't already present in the **Workbooks** collection.




```
On Error Resume Next ' turn off error checking 
Set wbMyAddin = Workbooks(AddIns("My Addin").Name) 
lastError = Err 
On Error Goto 0 ' restore error checking 
If lastError <> 0 Then 
 ' the add-in workbook isn't currently open. Manually open it. 
 Set wbMyAddin = Workbooks.Open(AddIns("My Addin").FullName) 
End If
```


## Properties



|**Name**|
|:-----|
|[Application](addin-application-property-excel.md)|
|[CLSID](addin-clsid-property-excel.md)|
|[Creator](addin-creator-property-excel.md)|
|[FullName](addin-fullname-property-excel.md)|
|[Installed](addin-installed-property-excel.md)|
|[IsOpen](addin-isopen-property-excel.md)|
|[Name](addin-name-property-excel.md)|
|[Parent](addin-parent-property-excel.md)|
|[Path](addin-path-property-excel.md)|
|[progID](addin-progid-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
