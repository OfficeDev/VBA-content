---
title: Hyperlink.CreateNewDocument Method (Access)
keywords: vbaac10.chm10121
f1_keywords:
- vbaac10.chm10121
ms.prod: access
api_name:
- Access.Hyperlink.CreateNewDocument
ms.assetid: bd0f0728-d2de-1b2b-529b-e3e9db41b660
ms.date: 06/08/2017
---


# Hyperlink.CreateNewDocument Method (Access)

You can use the  **CreateNewDocument** method to create a new document associated with a specified hyperlink.


## Syntax

 _expression_. **CreateNewDocument**( ** _FileName_**, ** _EditNow_**, ** _Overwrite_** )

 _expression_ A variable that represents a **Hyperlink** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**| The name and path of the document. The type of document format you want used can be determine by the extension used with the filename. to output the data. You can create HTML (*.htm), Microsoft Active Server Pages (*.asp), Microsoft Excel (*.xls), Microsoft IIS (*.htx, *.idc), MS-DOS Text (*.txt), or Rich Text Format (*.rtf). Modules can be output only to MS-DOS text format. Microsoft Internet Information Server and Microsoft Active Server formats are available only for tables, queries, and forms.|
| _EditNow_|Required|**Boolean**|**True** opens the document in design view and **False** stores the new document in the specified database directory. The default is **True**.|
| _Overwrite_|Required|**Boolean**|**True** overwrites an existing document if the _filename_ argument identifies an existing document and **False** requires that the filename argument specifies a new filename. The default is **False**.|

### Return Value

Nothing


## Remarks

The  **CreateNewDocument** method provides a way to programmatically create a document associated with a hyperlink within a control.


## Example

The following example utilizes a hyperlink control's  **Click** event. This event creates a new file named "Report.txt" when the user clicks the hyperlink control named "GenerateReport" on a form. The new file opened for editing. If a file named "Report.txt" already exists on drive C, it is replaced with this new file.


```vb
Private Sub GenerateReport_Click() 
 ActiveControl.Hyperlink.CreateNewDocument _ 
 "C:\Report.txt", EditNow:=True, Overwrite:=True 
End Sub
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-access.md)

