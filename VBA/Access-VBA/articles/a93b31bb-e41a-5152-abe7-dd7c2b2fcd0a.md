
# Append Method (ADOX Procedures)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

 **In this article**
[Syntax](#sectionSection1)
[Parameters](#sectionSection2)
[Remarks](#sectionSection3)



Adds a new [Procedure](d5fcf0fe-f59f-e114-dc11-515f11c2a2c1.md) object to the[Procedures](e1ca53ad-1213-b514-e015-e18c2ab15e23.md) collection.

## Syntax
<a name="sectionSection1"> </a>

 _Procedures_. **Append** _Name_, _Command_


## Parameters
<a name="sectionSection2"> </a>


-  _Name_
    
- A  **String** value that specifies the name of the procedure to create and append.
    
-  _Command_
    
- An ADO [Command](64f4ef03-f858-c004-b891-0c96d13a5e6e.md) object that represents the procedure to create and append.
    

## Remarks
<a name="sectionSection3"> </a>

Creates a new procedure in the data source with the name and attributes specified in the  **Command** object.

If the command text that the user specifies represents a view rather than a procedure, the behavior is dependent upon the provider being used.  **Append** will fail if the provider does not support persisting commands.


 **Note**  When using the OLE DB Provider for Microsoft Jet, the  **Procedures** collection **Append** method will allow you to specify a **View** rather than a **Procedure** in the _Command_ parameter. The **View** will be added to the data source and will be added to the **Procedures** collection. After the **Append**, if the **Procedures** and **Views** collections are refreshed, the **View** will no longer be in the **Procedures** collection and will appear in the **Views** collection.

