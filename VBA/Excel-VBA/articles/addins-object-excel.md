---
title: AddIns Object (Excel)
keywords: vbaxl10.chm186072
f1_keywords:
- vbaxl10.chm186072
ms.prod: excel
api_name:
- Excel.AddIns
ms.assetid: 2e9d9a1f-8833-beb3-757c-a5b26568f5fb
ms.date: 06/08/2017
---


# AddIns Object (Excel)

A collection of  **[AddIn](addin-object-excel.md)** objects that represents all the add-ins available to Microsoft Excel, regardless of whether they're installed.


## Remarks

This list corresponds to the list of add-ins displayed in the  **Add-Ins** dialog box.


## Example

Use the  **[AddIns](addins-application-property-excel.md)** property to return the **AddIns** collection. The following example creates a list that contains the names and installed states of all the available add-ins.


```vb
Sub DisplayAddIns() 
 Worksheets("Sheet1").Activate 
 rw = 1 
 For Each ad In Application.AddIns 
 Worksheets("Sheet1").Cells(rw, 1) = ad.Name 
 Worksheets("Sheet1").Cells(rw, 2) = ad.Installed 
 rw = rw + 1 
 Next 
End Sub
```

Use the  **[Add](addins-add-method-excel.md)** method to add an add-in to the list of available add-ins. The **Add** method adds an add-in to the list but doesn't install the add-in. Set the[Installed](addin-installed-property-excel.md) property of the add-in to **True** to install the add-in. To install an add-in that doesn't appear in the list of available add-ins, you must first use the **Add** method and then set the **Installed** property. This can be done in a single step, as shown in the following example (note that you use the name of the add-in, not its title, with the **Add** method).




```vb
AddIns.Add("generic.xll").Installed = True
```

Use  **AddIns** ( _index_ ) where _index_ is the add-in title or index number to return a single **AddIn** object. The following example installs the Analysis Toolpak add-in.

Don't confuse the add-in title, which appears in the  **Add-Ins** dialog box, with the add-in name, which is the file name of the add-in. You must spell the add-in title exactly as it's spelled in the **Add-Ins** dialog box, but the capitalization doesn't have to match.




```vb
AddIns("analysis toolpak").Installed = True
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


