
# Report.MoveLayout Property (Access)

The  **MoveLayout** property specifies whether Microsoft Access should move to the next printing location on the page. Read/write **Boolean**.


## Syntax

 _expression_. **MoveLayout**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **MoveLayout** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|**True**|(Default) The section's  **Left** and **Top** properties are advanced to the next print location.|
|**False**|The section's  **Left** and **Top** properties are unchanged.|
To set this property, specify an [event procedure](3fa3677b-a779-3bc7-0f0f-827c252b3292.md)for a section's  **[OnFormat](061652a9-0253-8dc2-a8c0-02daa40d132d.md)** property.

Microsoft Access sets this property to  **True** before each section's **Format** event.


## Example

The following example sets the  **MoveLayout** property for the "Purchase Order" report to its default setting.


```vb
Reports("Purchase Order").MoveLayout = True 

```


## See also


#### Concepts


[Report Object](6f77c1b4-a9ce-7caa-204c-fe0755c6f9df.md)
#### Other resources


[Report Object Members](73370a33-1ca0-da4d-9e36-88011bc2b93e.md)
