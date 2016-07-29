
# TableDef Members (DAO)

 **Last modified:** March 09, 2015

 _ **Applies to:** Access 2013 | Access 2016_

A  **TableDef** object represents the stored definition of a base table or a linked table (Microsoft Access workspaces only).


## Methods



|**Name**|**Description**|
|:-----|:-----|
|**[CreateField](a83d797f-ea42-4a07-dd9e-b254755f0891.md)**|Creates a new  **[Field](47282ce2-9b49-ccf9-ad37-c4bb25cfd037.md)** object (Microsoft Access workspaces only). .|
|**[CreateIndex](857b25c1-01fa-b926-0c74-7105e71b7505.md)**|Creates a new  **[Index](92c32cad-ec8a-1243-1d18-83f50b269ecb.md)** object (Microsoft Access workspaces only). .|
|**[CreateProperty](8a92cc64-414e-f33c-1c3e-d1b62c1688c2.md)**|Creates a new user-defined  **[Property](a1ecb0db-bb93-a7b5-23c3-0b73f275dfe0.md)** object (Microsoft Access workspaces only).|
|**[OpenRecordset](f4c9c89c-3348-d3c9-ce76-dd11e5ee11a7.md)**|Creates a new  **[Recordset](9774232c-e6da-175b-fc7f-ed2ab7908fa0.md)** object and appends it to the **Recordsets** collection.|
|**[RefreshLink](9f0059c6-3b7b-57e3-7527-ef674ad9417d.md)**|Updates the connection information for a linked table (Microsoft Access workspaces only).|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|**[Attributes](d01588c3-e94e-06bd-6568-974873411f2d.md)**|Sets or returns a value that indicates one or more characteristics of a  **TableDef** object. Read/write **Long**.|
|**[ConflictTable](0db8b975-eb6d-19c6-cfb7-6ce01230ebe4.md)**|Returns the name of a conflict table containing the database records that conflicted during the synchronization of two replicas (Microsoft Access workspaces only). Read-only  **String**.|
|**[Connect](4fbb324c-a358-8fad-60f2-fb8005cf74d9.md)**|Sets or returns a value that provides information about a linked table. Read/write  **String**.|
|**[DateCreated](fedd28e9-41a4-db7f-9ba9-6ada350d594a.md)**|Returns the date and time that an object was created (Microsoft Access workspaces only). Read-only  **Variant**.|
|**[Fields](ca85be33-c872-309d-b1f0-d1ffb6951547.md)**|Returns a  **Fields** collection that represents all stored **Field** objects for the specified object. Read-only.|
|**[Indexes](b168ff75-0a5f-2bc0-9180-2add520a12c6.md)**|Returns an  **Indexes** collection that contains all of the stored **Index** objects for the specified table. Read-only.|
|**[LastUpdated](fafe54e2-2cf0-5874-92b9-6e20a65e77ef.md)**|Returns the date and time of the most recent change made to an object. Read-only  **Variant**.|
|**[Name](66b751ee-cf8a-a1f2-c646-6124e5f18cd0.md)**|Returns or sets the name of the specified object. Read/write  **String**.|
|**[Properties](e6eefc5f-498c-77c1-79e1-e4d0b8cc2133.md)**|Returns the  **[Properties](cd07184a-a261-29c9-542f-bc2eff6f4af6.md)** collection of the specified object. Read-only.|
|**[RecordCount](f8804244-0134-fc1f-1f5f-4971afe17974.md)**|Returns the total number of records in a  **[TableDef](715146b6-c62a-abff-28ee-e6bbe3c08adf.md)** object. Read-only **Long**.|
|**[ReplicaFilter](f44273de-2b6a-750f-cb7c-12c3ac2da503.md)**|Sets or returns a value on a  **[TableDef](715146b6-c62a-abff-28ee-e6bbe3c08adf.md)** object within a partial replica that indicates which subset of records is replicated to that table from a full replica. (Microsoft Access workspaces only).|
|**[SourceTableName](3c02f5f6-70ae-39ec-0984-8d6b81992418.md)**|Sets or returns a value that specifies the name of a linked table or the name of a base table (Microsoft Access workspaces only).|
|**[Updatable](0b1ae7e5-416d-06f0-5d74-989c6db67ff2.md)**|Returns a value that indicates whether you can change a DAO object. Read-only  **Boolean**.|
|**[ValidationRule](7dcd6f2c-45bc-a50b-727d-589371d5803f.md)**|Sets or returns a value that validates the data in a field as it's changed or added to a table (Microsoft Access workspaces only).Read/write  **String**.|
|**[ValidationText](9f38616a-41ee-cbd1-9e29-da436b258e08.md)**|Sets or returns a value that specifies the text of the message that your application displays if the value of a  **Field** object doesn't satisfy the validation rule specified by the **ValidationRule** property setting (Microsoft Access workspaces only). Read/write **String**.|
