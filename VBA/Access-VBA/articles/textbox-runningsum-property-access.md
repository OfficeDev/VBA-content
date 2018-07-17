---
title: TextBox.RunningSum Property (Access)
keywords: vbaac10.chm11070
f1_keywords:
- vbaac10.chm11070
ms.prod: access
api_name:
- Access.TextBox.RunningSum
ms.assetid: 8918a58c-8c07-84dc-f43c-2486d54cd677
ms.date: 06/08/2017
---


# TextBox.RunningSum Property (Access)

You can use the  **RunningSum** property to calculate record-by-record or group-by-group totals in a report. Read/write **Byte**.


## Syntax

 _expression_. **RunningSum**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

The  **RunningSum** property specifies whether a text box on a report displays a running total and lets you set the range over which values are accumulated. For example, you can group data by month and show the sum of each month's sales in the group footer. You can show the running sum of accumulated sales over the entire report (sales for January in the January footer, sales for January plus February in the February footer, and so on) by adding a text box to the footer that shows the sum of sales and setting its **RunningSum** property to Over All.

The  **RunningSum** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|No|0|(Default) The text box displays the data from the underlying field for the current record.|
|Over Group|1|The text box displays a running sum of values in the same group level. The value accumulates until another group level section is encountered.|
|Over All|2|The text box displays a running sum of values in the same group level. The value accumulates until the end of the report.|

 **Note**  The  **RunningSum** property applies only to a text box on a report.

Place the text box in the Detail section to calculate a record-by-record total. For example, to number the records appearing in a detail section of a report, set the  **ControlSource** property for the text box to "=1", and set the **RunningSum** property to Over Group.

Place the text box in a group header or group footer to calculate a group-by-group total.

You can have up to 10 nested group levels in a report.


## Example

The following example sets the  **RunningSum** property for a text box named SalesTotal to 2 (Over All):


```vb
Reports!rptSales!SalesTotal.RunningSum = 2
```


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

