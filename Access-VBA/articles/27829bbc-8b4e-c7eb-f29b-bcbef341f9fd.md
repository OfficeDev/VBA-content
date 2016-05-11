
# Field2 Members (DAO)

 **Last modified:** March 09, 2015

 _ **Applies to:** Access 2013 | Access 2016_

A  **Field2** object represents a column of data with a common data type and a common set of properties.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|**[AppendChunk](540cd02d-1fc6-81d1-ac08-1e3df72a7208.md)**|Appends data from a string expression to a Memo or Long Binary  **Field2** object in a **[Recordset](9774232c-e6da-175b-fc7f-ed2ab7908fa0.md)**.|
|**[CreateProperty](bdbd6bec-216f-138e-78df-9c3221692aa4.md)**|Creates a new user-defined  **[Property](a1ecb0db-bb93-a7b5-23c3-0b73f275dfe0.md)** object (Microsoft Access workspaces only).|
|**[GetChunk](5d3a66c0-8216-d701-0a91-b79fbbc822b8.md)**|Returns all or a portion of the contents of a  **Memo** or **Long Binary** **Field2** object in the **[Fields](4be3ba07-20c1-d958-c1b8-7dd8b4731f60.md)** collection of a **[Recordset](9774232c-e6da-175b-fc7f-ed2ab7908fa0.md)** object.|
|**[LoadFromFile](8ffe4636-d4da-0579-f4b5-14f423647562.md)**|Loads the specified file from disk.|
|**[SaveToFile](250f9596-1a03-471d-96f9-718cd57dc94f.md)**|Saves an attachment to disk. .|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|**[AllowZeroLength](d3795634-527f-b4c5-b606-50f9945cac12.md)**|Sets or returns a value that indicates whether a zero-length string ("") is a valid setting for the  **[Value](6c0f9a8d-f51a-b8cf-8830-f8d960a1d08c.md)** property of the **Field2** object with a Text or Memo data type (Microsoft Access workspaces only).|
|**[AppendOnly](4427f3af-6393-0f1c-ecac-017112022583.md)**|Gets or sets a  **Boolean** that indicates whether the spcified field is set to append new values to the existing contents of the field as they are added. Read/write.|
|**[Attributes](08ae9b6b-21e4-9b7e-0852-cfc6639027a7.md)**|Sets or returns a value that indicates one or more characteristics of a  **Field2** object. Read/write **Long**.|
|**[CollatingOrder](cb1d6fc9-a2a6-54c2-abf5-48b609e38738.md)**|Returns a value that specifies the sequence of the sort order in text for string comparison or sorting (Microsoft Access workspaces only). Read-only  **Long**.|
|**[ComplexType](9b4ebabf-22de-0ab8-73ea-10c496eedf97.md)**|Returns a  **[ComplexType](fc9bdebe-e432-e530-6b1f-8680b9dfe870.md)** object that represents a multi-valued field. Read-only.|
|**[DataUpdatable](e6619c4e-26b1-777b-f0de-78fed3dbc890.md)**|Returns a value that indicates whether the data in the field represented by a  **Field2** object is updatable.|
|**[DefaultValue](709c9580-520e-46ce-7d70-e409872184bb.md)**|Sets or returns the default value of a  **Field2** object. For a **Field2** object not yet appended to the **[Fields](4be3ba07-20c1-d958-c1b8-7dd8b4731f60.md)** collection, this property is read/write (Microsoft Access workspaces only).|
|**[Expression](8ae9db2c-7460-5bfc-0dc4-3f87e5ab30ff.md)**|Read/write|
|**[FieldSize](d609801d-7761-663f-2840-de5923bb120c.md)**|Returns the number of bytes used in the database (rather than in memory) of a Memo or Long Binary  **Field2** object in the **[Fields](4be3ba07-20c1-d958-c1b8-7dd8b4731f60.md)** collection of a **[Recordset](9774232c-e6da-175b-fc7f-ed2ab7908fa0.md)** object.|
|**[ForeignName](76da233a-efb4-63cd-a2a2-d18d9e2fb2fb.md)**|Sets or returns a value that specifies the name of the  **Field2** object in a foreign table that corresponds to a field in a primary table for a relationship (Microsoft Access workspaces only).|
|**[IsComplex](ffc90e6e-e3ee-4f9b-ca6b-615199300d45.md)**|Returns  **Boolean** that indicates whether the specified field is a multi-valued data type. Read-only.|
|**[Name](6f84ca11-4e7c-9573-5261-b67b91ba30dc.md)**|Returns or sets the name of the specified object. Read/write  **String** if the object has not been appended to a collection. Read-only **String** if the object has been appended to a collection.|
|**[OrdinalPosition](55d89611-ad07-990d-fc33-f81d59472430.md)**|Sets or returns the relative position of a  **Field2** object within a **[Fields](4be3ba07-20c1-d958-c1b8-7dd8b4731f60.md)** collection. .|
|**[OriginalValue](10fed55e-c938-2ae6-8fd2-996745a63da3.md)**|
 **Note**  ODBCDirect workspaces are not supported in Microsoft Access 2013. Use ADO if you want to access external data sources without using the Microsoft Access database engine.

Returns the value of a  **Field2** in the database that existed when the last batch update began (ODBCDirect workspaces only).|
|**[Properties](a365c2ef-c9b5-d765-e609-2e070c66de55.md)**|Returns the  **[Properties](cd07184a-a261-29c9-542f-bc2eff6f4af6.md)** collection of the specified object. Read-only.|
|**[Required](7d14dfd7-a50d-6044-469e-1511c74c148d.md)**|Sets or returns a value that indicates whether a  **Field2** object requires a non-Null value.|
|**[Size](e252759a-cea9-25cb-667d-80b422fbf97e.md)**|Sets or returns a value that indicates the maximum size, in bytes, of a  **Field2** object.|
|**[SourceField](f89146c1-d4a4-1129-636a-c22cf7921a4e.md)**|Returns a value that indicates the name of the field that is the original source of the data for a  **Field2** object. Read-only **String**.|
|**[SourceTable](24d2fdda-d924-d95e-8458-063dd1d36d5d.md)**|Returns a value that indicates the name of the table that is the original source of the data for a  **Field2** object. Read-only **String**.|
|**[Type](057d6ec9-b72c-cee6-005a-6d916e3dda29.md)**|Sets or returns a value that indicates the operational type or data type of an object. Read/write  **Integer**.|
|**[ValidateOnSet](07612730-8dad-4ef0-b19b-f76845973fc3.md)**|Sets or returns a value that specifies whether or not the value of a  **Field2** object is immediately validated when the object's **Value** property is set (Microsoft Access workspaces only).|
|**[ValidationRule](5464d2de-f3d7-5d6b-4fc5-66df6a5540cb.md)**|Sets or returns a value that validates the data in a field as it's changed or added to a table (Microsoft Access workspaces only). Read/write  **String**.|
|**[ValidationText](6128f66c-3891-ae4d-d12d-354a79a9c05e.md)**|Sets or returns a value that specifies the text of the message that your application displays if the value of a  **Field2** object doesn't satisfy the validation rule specified by the **ValidationRule** property setting (Microsoft Access workspaces only). Read/write **String**.|
|**[Value](6ead6ba8-1613-99c7-7968-56f5b81b2385.md)**|Sets or returns the value of an object. Read/write  **Variant**.|
|**[VisibleValue](1e023a1a-fd49-7570-42bd-2f4c06ac5e5e.md)**|
 **Note**  ODBCDirect workspaces are not supported in Microsoft Access 2013. Use ADO if you want to access external data sources without using the Microsoft Access database engine.

Returns a value currently in the database that is newer than the  **OriginalValue** property as determined by a batch update conflict (ODBCDirect workspaces only).|
