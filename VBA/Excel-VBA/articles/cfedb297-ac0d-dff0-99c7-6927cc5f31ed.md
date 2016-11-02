
# Name Object (Excel)

Represents a defined name for a range of cells. Names can be either built-in names — such as Database, Print_Area, and Auto_Open — or custom names.


## Remarks

 **Application, Workbook, and Worksheet Objects**

The  **Name** object is a member of the **[Names](ffecf89d-7bae-c470-8e37-608857a9de2a.md)** collection for the **[Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)**, **[Workbook](8c00aa60-c974-eed3-0812-3c9625eb0d4c.md)**, and **[Worksheet](182b705e-854a-81cc-a4b0-59b942de55ae.md)** objects. Use **[Names](26be56ec-ea12-1600-602a-eb338d4a5a8b.md)** ( _index_ ), where _index_ is the name index number or defined name, to return a single **Name** object.

The index number indicates the position of the name within the collection. Names are placed in alphabetic order, from a to z, and are not case-sensitive.

 **Range Objects**

Although a  **[Range](b8207778-0dcc-4570-1234-f130532cc8cd.md)** object can have more than one name, there's no **Names** collection for the **Range** object. Use **[Name](39d1a326-e123-443c-29c0-453f7b4a8760.md)** with a **Range** object to return the first name from the list of names (sorted alphabetically) assigned to the range. The following example sets the **[Visible](48860564-6079-932e-2cae-0802235be61e.md)** property for the first name assigned to cells A1:B1 on worksheet one.


## Example

The following example displays the cell reference for the first name in the application collection.


```
MsgBox Names(1).RefersTo
```

The following example deletes the name "mySortRange" from the active workbook.




```
ActiveWorkbook.Names("mySortRange").Delete
```

Use the  **Name** property to return or set the text of the name itself. The following example changes the name of the first **Name** object in the active workbook.




```
Names(1).Name = "stock_values"
```

The following example sets the  **Visible** property for the first name assigned to cells A1:B1 on worksheet one.




```
Worksheets(1).Range("a1:b1").Name.Visible = False
```


## Methods



|**Name**|
|:-----|
|[Delete](429a5d17-8f34-9a04-d744-66ce1e9e39a7.md)|

## Properties



|**Name**|
|:-----|
|[Application](e8272a17-5ad8-b63f-3b30-7abd49434d98.md)|
|[Category](01892c7b-a42e-e4b3-6ddd-27ace1c51aae.md)|
|[CategoryLocal](5f80e0a4-e12d-a85d-69a1-979652f62ac3.md)|
|[Comment](7d2e9c31-4c81-f1ae-1c8b-a476c2bc0d7f.md)|
|[Creator](90c6fe07-e941-269f-71bf-e9dc6a982629.md)|
|[Index](b7c5c593-80d3-d36a-ec68-7733bbb7e5a8.md)|
|[MacroType](46f02cb6-56c3-7b0e-27a4-db356802abe6.md)|
|[Name](eeebe875-b60d-7abe-df4e-8b56476b6b64.md)|
|[NameLocal](7a98f361-077f-30fc-b754-4070e526f7bc.md)|
|[Parent](83d46498-bf9c-6285-189b-47f6e8cd41ee.md)|
|[RefersTo](8093e14c-0461-5e49-ef71-16c683044a63.md)|
|[RefersToLocal](e079e8c9-44f9-494e-97aa-2a38c0ec157b.md)|
|[RefersToR1C1](6661dc25-44cd-ac43-9347-93ed7583c9b1.md)|
|[RefersToR1C1Local](314b8764-5f5c-9a2f-87a7-54637de59bbd.md)|
|[RefersToRange](81c0e2fe-8ce6-0df9-9ffa-0931b87487e7.md)|
|[ShortcutKey](ff763568-4c18-9414-45a7-bcf75b597261.md)|
|[ValidWorkbookParameter](fd8bef70-af4f-af01-1956-24b50ea210be.md)|
|[Value](26732c54-3519-885d-e40d-69c6b1795318.md)|
|[Visible](078a949c-ff27-c62d-10b0-7d83b190da13.md)|
|[WorkbookParameter](1a7983fc-9020-fb72-21b1-822d19802c31.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)