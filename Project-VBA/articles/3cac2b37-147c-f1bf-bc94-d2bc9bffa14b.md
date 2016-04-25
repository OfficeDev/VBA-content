
# Application.GanttBarStyleDelete Method (Project)

Deletes a Gantt bar style from the active Gantt Chart.


## Syntax

 _expression_. **GanttBarStyleDelete**( ** _Item_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required|**String**|**String**. The name or row number of the Gantt bar to delete from the **Bar Styles** dialog box.|

### Return Value

 **Boolean**


## Remarks

To manually show the  **Bar Styles** dialog box, click the **Format** tab under the **Gantt Chart Tools** tab. In the **Bar Styles** group, click **Bar Styles** in the **Format** drop-down list. The **Bar Styles** dialog box can contain up to 200 style entries.


## Example

The following command deletes style number 41 in the  **Bar Styles** dialog box.


```
GanttBarStyleDelete Item:="41"
```

