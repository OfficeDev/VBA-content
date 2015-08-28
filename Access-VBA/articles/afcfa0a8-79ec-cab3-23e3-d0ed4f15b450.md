
# AllStoredProcedures.Application Property (Access)

 **Last modified:** July 28, 2015

You can use the  **Application** property to access the active Microsoft Access ** [Application](aefb0713-97e6-e2c7-e530-8fd2e1316a55.md)**object and its related properties. Read-only  **Application** object.

## Syntax

 _expression_. **Application**

 _expression_A variable that represents an  **AllStoredProcedures** object.


## Remarks

The  **Application** property is set by Microsoft Access and is read-only in all views.

Each Microsoft Access object has an  **Application** property that returns the current **Application** object. You can use this property to access any of the object's properties. For example, you could refer to the menu bar for the **Application** object from the current form by using the following syntax:




```
Me.Application.MenuBar 

```


## See also


#### Concepts


 [AllStoredProcedures Collection](896f4c2c-273c-2849-0f06-d75fa515c44a.md)
#### Other resources


 [AllStoredProcedures Object Members](36a5b27f-9a95-d7e4-bca1-7de9252893a4.md)
