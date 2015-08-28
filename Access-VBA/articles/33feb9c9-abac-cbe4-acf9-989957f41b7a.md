
# AccessObject.GetDependencyInfo Method (Access)

 **Last modified:** July 28, 2015

 Returns a ** [DependencyInfo](46ccdc3f-0101-5d81-8c01-ac37f139a2bc.md)** object that represents the database objects that are dependent upon the specified object.

## Syntax

 _expression_. **GetDependencyInfo**

 _expression_A variable that represents an  **AccessObject** object.


### Return Value

DependencyInfo


## Remarks

This method will return a run-time error if any of the following conditions are true:


- The  **Track name AutoCorrect info** setting ( **Tools** menu, **Options** dialog box, **General** tab) is disabled. You can use the following code to enable the **Track name AutoCorrect info** setting and update the dependency information for all of the objects in the database: `Application.SetOption "Track Name AutoCorrect Info", 1`
    
- You have insufficient permissions to check the dependency information for the specified  **AccessObject** object.
    
- This method is being called from an Access project (.adp).
    


Access does not search Visual Basic for Applications (VBA) code, macros, or data access pages for dependencies.


## See also


#### Concepts


 [AccessObject Object](8a770b33-5bff-120a-6707-ca214ee5ced3.md)
#### Other resources


 [AccessObject Object Members](78aaacb1-c0d3-d809-088d-d543ecd71de3.md)
