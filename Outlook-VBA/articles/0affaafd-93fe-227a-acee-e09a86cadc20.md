
# Table Object (Outlook)

 **Last modified:** July 28, 2015

Represents a set of item data from a  ** [Folder](3cf6cda8-6d70-666e-2643-9d9c5b9cacfc.md)** or ** [Search](226a5d49-3caf-90dd-725c-265404d1939f.md)** object, with items as rows of the table and properties as columns of the table.

## Remarks

The  **Table** represents a read-only dynamic rowset of data in a **Folder** or **Search** object. You can use ** [Folder.GetTable](08d184cb-0c41-01b1-abc5-305476380f8b.md)** or ** [Search.GetTable](3aba6b77-73a3-9620-9c18-b2e03c7b63bc.md)** to obtain a **Table** object that represents a set of items in a folder or search folder. If the **Table** object is obtained from **Folder.GetTable**, you can further specify a filter (in  ** [Table.Restrict](ecdd30f6-e12c-8025-3ded-592d2fad2bb8.md)**) to obtain a subset of the items in the folder. If you do not specify any filter, you will obtain all the items in the folder. 

By default, each item in the returned  **Table** contains only a default subset of its properties. You can regard each row of a **Table** as an item in the folder, each column as a property of the item, and the **Table** as an in-memory lightweight rowset that allows fast enumeration and filtering of items in the folder. Although additions and deletions of the underlying folder are reflected by the rows in the **Table**, the  **Table** does not support any events for adding, changing, and removing of rows. If you require a writeable object from the **Table** row, obtain the Entry ID for that row from the default EntryID column in the **Table** and then use the ** [GetItemFromID](f2abff80-4c04-998b-654b-28600424a16f.md)** method of the ** [NameSpace](f0dcaa19-07f5-5d42-a3bf-2e42b7885644.md)** object to obtain a full item, such as a ** [MailItem](14197346-05d2-0250-fa4c-4a6b07daf25f.md)** or ** [ContactItem](8e32093c-a678-f1fd-3f35-c2d8994d166f.md)**, that supports read-write operations. For more information on default columns in a  **Table**, see  [Default Properties Displayed in a Table Object](649c64f3-2d1e-23f1-bf13-3368da79e62b.md).

 For more information on the **Table** object, see [Enumerating, Searching, and Filtering Items in a Folder](d786d292-7a0e-0e1a-e132-affbfde37744.md).


## Example

The following code sample illustrates how the  **Table** object can return a filtered set of items based on their **LastModificationTime** property. It also shows how to list the default properties as well as specific properties of the items.


```
Sub DemoTable() 
 
 'Declarations 
 
 Dim Filter As String 
 
 Dim oRow As Outlook.Row 
 
 Dim oTable As Outlook.Table 
 
 Dim oFolder As Outlook.Folder 
 
 
 
 'Get a Folder object for the Inbox 
 
 Set oFolder = Application.Session.GetDefaultFolder(olFolderInbox) 
 
 
 
 'Define Filter to obtain items last modified after May 1, 2005 
 
 Filter = "[LastModificationTime] > '5/1/2005'" 
 
 'Restrict with Filter 
 
 Set oTable = oFolder.GetTable(Filter) 
 
 
 
 'Remove all columns in the default column set 
 
 oTable.Columns.RemoveAll 
 
 'Specify desired properties 
 
 With oTable.Columns 
 
 .Add ("Subject") 
 
 .Add ("LastModificationTime") 
 
 'PR_ATTR_HIDDEN referenced by the MAPI proptag namespace 
 
 .Add ("http://schemas.microsoft.com/mapi/proptag/0x10F4000B") 
 
 End With 
 
 
 
 'Enumerate the table using test for EndOfTable 
 
 Do Until (oTable.EndOfTable) 
 
 Set oRow = oTable.GetNextRow() 
 
 Debug.Print (oRow("Subject")) 
 
 Debug.Print (oRow("LastModificationTime")) 
 
 Debug.Print (oRow("http://schemas.microsoft.com/mapi/proptag/0x10F4000B")) 
 
 Loop 
 
End Sub
```


## See also


#### Concepts


 [Outlook Object Model Reference](73221b13-d8d8-99b8-3394-b95dbbfd5ddc.md)
#### Other resources


 [Table Object Members](bd9db35d-0738-22cf-a936-425d5a0ead87.md)
