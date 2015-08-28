
# TempVars Object (Access)

 **Last modified:** July 28, 2015

Represents the collection of  ** [TempVar](4a0429e6-bcfa-7a8b-7030-6e88c2f1a71d.md)** objects.

## Remarks

Use the  ** [Add](836e449c-35ff-4089-857a-403c9fc97592.md)** method or the [SetTempVar](http://msdn.microsoft.com/library/9c3b7bee-02c5-efbf-1276-4c4a1f7802d9%28Office.15%29.aspx) macro action to create a **TempVar** object.

Use the  ** [Remove](a9ab9ff2-5bfc-d001-f5eb-9929907bc1b2.md)** method or the [RemoveTempVar](http://msdn.microsoft.com/library/409fd836-4a53-cefd-4264-8cee0fa8ac52%28Office.15%29.aspx) macro action to delete a **TempVar** object from the **TempVars** collection.

Use the  ** [RemoveAll](1b278bda-9f28-8fd7-0408-3a2a4d3e1a74.md)** method or [RemoveAllTempVars](http://msdn.microsoft.com/library/409fd836-4a53-cefd-4264-8cee0fa8ac52%28Office.15%29.aspx) macro action to delete all **TempVar** objects from the **TempVars** collection.

The  **TempVars** collection can store up to 255 **TempVar** objects. If you do not remove a **TempVar** object, it will remain in memory until you close the database. It is a good practice to remove **TempVar** object variables when you are finished using them.

To refer to a  **TempVar** object in a collection by its ordinal number or by its **Name** property setting, use the following syntax form:


-  **TempVar**![name]
    

## See also


#### Concepts


 [Access Object Model Reference](2de134a4-6c5c-d2a3-8377-f4dd973ba650.md)
#### Other resources


 [TempVars Object Members](5c83c870-c66c-8fd9-0ac6-06766b14a6fc.md)
