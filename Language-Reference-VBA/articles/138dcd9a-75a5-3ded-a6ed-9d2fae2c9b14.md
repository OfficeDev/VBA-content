
# Can't Get or Put user-defined type containing object reference

 **Last modified:** July 28, 2015

An object reference is temporary and can easily become invalid between closing and opening a file. This error has the following cause and solution:




- The  [variable](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) in your **Get** or **Put** statement contains, or is declared to contain, a reference to an object.
    
    If the variable is an object reference you can't use it with  **Get** and **Put** statements. To place the value of some or all of the object's [properties](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) in the file, each property must be individually specified.
    
- The  [user-defined type](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) variable in your **Get** or **Put** statement contains an element that is an object reference.
    
    If the variable's  **Type** statement contains an element representing an object (for example, it is defined in a [class module](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md), has  [Object data type](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md), is a form or a control, and so on), remove it from the definition, or define a new type for use with the  **Get** and **Put** statements that has no **Object** type element in its definition.
    
    If you have elements in the user-defined type with  **Variant** type, make sure no object reference is assigned to that element. A **Variant** can accept such an assignment, but will cause this error if its user-defined type is used in a **Get** or **Put**.
    
    Note that you can use  **Input #**,  **Line Input #**,  **Print #**, or  **Write #** to write the default property of an object to disk.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).
