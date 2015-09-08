
# DoCmd.SearchForRecord Method (Access)

 **Last modified:** July 28, 2015

You can use the  **SearchForRecord** method to search for a specific record in a table, query, form, or report.

## Syntax

 _expression_. **SearchForRecord**( **_ObjectType_**,  **_ObjectName_**,  **_Record_**,  **_WhereCondition_**)

 _expression_A variable that represents a  **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ObjectType|Optional| **AcDataObjectType**|An  ** [AcDataObjectType](0e9f8481-ef01-2415-414a-64788c18e6ef.md)** constant that specifies the type of database object in which you are searching. The default value is **acActiveDataObject**.|
|ObjectName|Optional| **Variant**|The name of the database object that contains the record to search for.|
|Record|Optional| **AcRecord**|An  ** [AcRecord](39ece328-d461-9f4d-a3af-205ed3228929.md)** constant that specifies the starting point and direction of the search. The default value is **acFirst**.|
|WhereCondition|Optional| **Variant**|A string used to locate the record. It is like the WHERE clause in an SQL statement, but without the word WHERE.|

## Remarks




- In cases where more than one record matches the criteria in the WhereCondition argument, the following factors determine which record is found:
    
      - The Record argument setting.
    
  - The sort order of the records. For example, if the Record argument is set to  **acFirst**, changing the sort order of the records might change which record is found.
    
- The object specified in the ObjectName argument must be open before this action is run. Otherwise, an error occurs.
    
- If the criteria in the WhereCondition argument are not met, no error occurs and the focus remains on the current record.
    
- When searching for the previous or next record, the search does not "wrap" when it reaches the end of the data. If there are no further records that match the criteria, no error occurs and the focus remains on the current record. To confirm that a match was found, you can enter a condition for the next action, and make the condition the same as the criteria in the WhereCondition argument.
    
- The  **SearchForRecord** method is similar to the ** [FindRecord](dc48bc3d-5408-40a8-509b-e52b48b26187.md)** method, but **SearchForRecord** has more powerful search features. The **FindRecord** method is primarily used for finding strings, and it duplicates the functionality of the **Find** dialog box. The **SearchForRecord** method uses criteria that are more like those of a filter or an SQL query. The following list demonstrates some things you can do with the **SearchForRecord** method:
    
      - You can use complex criteria in the WhereCondition argument, such as  `Description = "Beverages" and CategoryID = 11`
    
    
    
  - You can refer to fields that are in the record source of a form or report but are not displayed on the form or report. In the preceding example, neither  `Description` nor `CategoryID` must be displayed on the form or report for the criteria to work.
    
  - You can use logical operators, such as  **&lt;**,  **&gt;**,  **AND**,  **OR**, and  **BETWEEN**. The  **FindRecord** method only matches strings that equal, start with, or contain the string being searched for.
    

## See also


#### Concepts


 [DoCmd Object](3ce44cca-9979-0a1e-9787-079a52ce528f.md)
#### Other resources


 [DoCmd Object Members](3e7ade9e-86e4-0751-188b-5d31c9101651.md)
