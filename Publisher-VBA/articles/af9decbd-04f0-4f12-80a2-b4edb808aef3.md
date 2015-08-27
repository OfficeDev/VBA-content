
# MailMergeDataField Properties (Publisher)

 **Last modified:** July 28, 2015


## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](6af180b7-99c6-85b3-bc7e-071bc655c4d8.md)|Used without an object qualifier, this property returns an  ** [Application](acfc7efb-e6a5-a89a-3aee-3cb4af2f3508.md)**object that represents the current instance of Publisher. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object. When used with an OLE Automation object, it returns the object's application.|
| [Creator](1f8f4da6-2d03-c3f4-3590-ebf82cffcd48.md)|Returns a  **Long** that represents the application in which the specified object was created. For example, if the object was created in Microsoft Publisher, this property returns the hexadecimal number 4D505542, which represents the string "MSPB." This value can also be represented by the constant.|
| [FieldType](9574f59b-a03f-ab0b-a2ac-085f31473f78.md)||
| [Index](f70d0266-0527-6871-632d-b45b617d75d4.md)|Returns a  **Long** that represents the position of a particular item in a specified collection. .|
| [IsMapped](4a053a2b-f6ca-37a7-4a1f-8690982188c2.md)|Indicates if the parent  **MailMergeDataField** object is mapped to a recipient field in the master data source (combined mail-merge recipient list). Read-only.|
| [MappedTo](067619e8-98fe-d0c2-2f50-96b50cf53de4.md)|Returns the name of the recipient field (column) in the master data source (combined mail-merge recipient list) that the parent  **MailMergeDataField** object is mapped to. Read-only.|
| [Name](7a2f4e1d-446c-707b-2375-8481e8f08cf5.md)|Returns a  **String** value indicating the name of the specified object. Read-only.|
| [Parent](cca35fe6-b959-0cfe-85de-347db2655c38.md)|Returns an object that represents the parent object of the specified object. For example, for a  ** [TextFrame](95e88f5a-b3dc-272e-7c1d-5282c97ae11e.md)** object, returns a ** [Shape](666cb7f0-62a8-f419-9838-007ef29506ee.md)** object representing the parent shape of the text frame. Read-only.|
| [Value](9ce1859b-72f0-c44e-0683-287c6e13b33c.md)|Returns a  **String** that represents the value of a mail merge data field record or a mapped data field. Read-only.|
