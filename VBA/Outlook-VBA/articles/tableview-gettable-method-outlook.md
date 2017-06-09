---
title: TableView.GetTable Method (Outlook)
keywords: vbaol11.chm3315
f1_keywords:
- vbaol11.chm3315
ms.prod: outlook
api_name:
- Outlook.TableView.GetTable
ms.assetid: 4f20a3cc-5ec9-a58b-8fcf-00e86f160493
ms.date: 06/08/2017
---


# TableView.GetTable Method (Outlook)

Returns a  **[Table](table-object-outlook.md)** object that represents all of the Microsoft Outlook items that are contained in a **[TableView](tableview-object-outlook.md)** object.


## Syntax

 _expression_ . **GetTable**

 _expression_ A variable that represents a **TableView** object.


### Return Value

A  **Table** whose rows represent items in the current table view.


## Remarks

The  **GetTable** method of the **TableView** object returns a table of items from one or more folders in the same store or spanning over multiple stores, in an aggregated view. For example, an aggregated view obtained by a search across all mail items by using Instant Search. This behavior differs from the **[GetTable](folder-gettable-method-outlook.md)** method of the **[Folder](folder-object-outlook.md)** object, which obtains a table object that contains items from the same folder.

 The parent **TableView** object must be based on the current folder of the active explorer, as indicated by the **[CurrentFolder](explorer-currentfolder-property-outlook.md)** property of the active **[Explorer](explorer-object-outlook.md)** object. If the folder is not a current folder of a visible explorer, or if the view of that folder, which is indicated by the **[Folder.CurrentView](folder-currentview-property-outlook.md)** property, is not a table view, Outlook returns an error.

The filter for the resultant table is set by the  **[Filter](tableview-filter-property-outlook.md)** property of the **TableView** object. If the **Filter** property of the **TableView** object is not empty, **GetTable** returns a **Table** object with rows that represent the filtered subset of items available in the view. If subsequently, the **[Table.Restrict](table-restrict-method-outlook.md)** method is called on the resultant table, applying the **Restrict** method is equivalent to a logical AND operation with the filter represented by **TableView.Filter** .

 **GetTable** returns a **Table** with the default column set. **GetTable** does not return a **Table** that contains columns for each field in the **[ViewFields](viewfields-object-outlook.md)** collection of the current view. For more information on the default column set of a table based on the folder type, see[Default Properties Displayed in a Table Object](http://msdn.microsoft.com/library/649c64f3-2d1e-23f1-bf13-3368da79e62b%28Office.15%29.aspx). To modify the default column set, use the  **[Add](columns-add-method-outlook.md)** , **[Remove](columns-remove-method-outlook.md)** , or **[RemoveAll](columns-removeall-method-outlook.md)** methods of the **[Columns](columns-object-outlook.md)** collection object. Properties that you cannot add to a table as columns are listed in[Unsupported Properties in a Table Object or Table Filter](http://msdn.microsoft.com/library/0e37f03f-7677-ca29-d0b2-8b45c026e5f1%28Office.15%29.aspx).

 The order of rows in the resultant table is not guaranteed to be the same as the order of items in the current view on which **GetTable** is based. For example, **GetTable** does not return a table with a row that represents a group-by header in the view. To sort the rows in the table returned from **GetTable** , use the **[Sort](table-sort-method-outlook.md)** method of the **Table** object.

The parent object of the  **Table** object returned by **GetTable** is the **TableView** object. The parent object of the **TableView** object is the **[Views](views-object-outlook.md)** collection, and the parent object of the **Views** collection is the **[Folder](folder-object-outlook.md)** object.


## Example

The following managed code is written in C#. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.

The following code example obtains a  **Table** object from the current view of the Inbox folder. The code sample sets the current folder of the active explorer to the Inbox, and then checks that the current view of the Inbox is a table view. After assuring these two conditions, the code sample calls the **TableView.GetTable** method and displays each item represented by each row in the returned **Table** .




```C#
private void DemoViewGetTable() 
{ 
 // Obtain the Inbox folder. 
 Outlook.Folder inbox = 
 Application.Session.GetDefaultFolder( 
 Outlook.OlDefaultFolders.olFolderInbox) 
 as Outlook.Folder; 
 
 // Set ActiveExplorer.CurrentFolder to Inbox. 
 // Inbox must be the current folder 
 // for TableView.GetTable to work correctly. 
 Application.ActiveExplorer().CurrentFolder = inbox; 
 
 // Ensure that the current view is a table view. 
 if (inbox.CurrentView.ViewType == 
 Outlook.OlViewType.olTableView) 
 { 
 Outlook.TableView view = 
 inbox.CurrentView as Outlook.TableView; 
 
 // No arguments are needed for View.GetTable. 
 Outlook.Table table = view.GetTable(); 
 
 Debug.WriteLine("View Count=" 
 + table.GetRowCount().ToString()); 
 while (!table.EndOfTable) 
 { 
 // First row in Table. 
 Outlook.Row nextRow = table.GetNextRow(); 
 Debug.WriteLine(nextRow["Subject"] 
 + " Modified: " 
 + nextRow["LastModificationTime"]); 
 } 
 } 
} 

```


## See also


#### Concepts


[TableView Object](tableview-object-outlook.md)
#### Other resources



[How to: Search and Obtain Items in an Aggregated View](http://msdn.microsoft.com/library/bd62f7b8-f110-ee0a-5930-877f14353a84%28Office.15%29.aspx)

