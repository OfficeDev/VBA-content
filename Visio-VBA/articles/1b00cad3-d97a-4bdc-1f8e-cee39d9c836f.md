
# Characters.AddField Method (Visio)

Replaces the text represented by a  **Characters** object with a new field of the category, code, and format you specify.


## Syntax

 _expression_ . **AddField**( **_Category_** , **_Code_** , **_Format_** )

 _expression_ A variable that represents a **Characters** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Category_|Required| **Integer**| **VisFieldCategories** . The category for the new field.|
| _Code_|Required| **Integer**| **VisFieldCodes** . The code for the new field.|
| _Format_|Required| **Integer**| **VisFieldFormats** . The format for the new field.|

### Return Value

Nothing


## Remarks

Using the  **AddField** method is similar to clicking **Field** on the **Insert** tab and inserting any of the following categories of fields in the text:


- Date/Time
    
- Document Info
    
- Geometry
    
- Object Info
    
- Page Info
    


To add a custom formula field, use the  **AddCustomField** method.

To specify language and calendary versions for Date/Time fields, use the  **AddFieldEx** method.

Constant values for  _Category_,  _Code_, and  _Format_ are declared by the Visio type library in **[VisFieldCategories](f10df918-5be3-e883-1da5-2a932fd1074f.md)** , **[VisFieldCodes](3bcc4aef-21c1-b152-47dc-74e6c58cd24e.md)** , and **[VisFieldFormats](ae671032-b96f-6652-f527-feb194a0261d.md)** respectively.

