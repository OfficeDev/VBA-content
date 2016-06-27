
# DefinedSize Property (ADO)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_



Indicates the data capacity of a [Field](1dbd535e-48ad-a5c8-a1b2-6776c1e3e19d.md) object.

## Return Value

Returns a  **Long** value that reflects the defined size of a field as a number of bytes.


## Remarks

Use the  **DefinedSize** property to determine the data capacity of a **Field** object.

The  **DefinedSize** and[ActualSize](020a414d-e6aa-5fb9-9b77-bd9d10124f8a.md) properties are different. For example, consider a **Field** object with a declared type of **adVarChar** and a **DefinedSize** property value of 50, containing a single character. The **ActualSize** property value it returns is the length in bytes of the single character.

