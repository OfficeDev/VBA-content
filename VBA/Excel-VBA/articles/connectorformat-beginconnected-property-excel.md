---
title: ConnectorFormat.BeginConnected Property (Excel)
keywords: vbaxl10.chm646077
f1_keywords:
- vbaxl10.chm646077
ms.prod: excel
api_name:
- Excel.ConnectorFormat.BeginConnected
ms.assetid: 2ebc4d15-e6f3-a0c9-056e-78004465c60c
ms.date: 06/08/2017
---


# ConnectorFormat.BeginConnected Property (Excel)

 **True** if the beginning of the specified connector is connected to a shape. Read-only **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **BeginConnected**

 _expression_ A variable that represents a **ConnectorFormat** object.


## Remarks





| **MsoTriState** can be one of these **MsoTriState** constants.|
| **msoCTrue**|
| **msoFalse**|
| **msoTriStateMixed**|
| **msoTriStateToggle**|
| **msoTrue** . The beginning of the specified connector is connected to a shape.|

## Example

If shape three on  `myDocument` is a connector whose beginning is connected to a shape, this example stores the connection site number in the variable `oldBeginConnSite`, stores a reference to the connected shape in the object variable  `oldBeginConnShape`, and then disconnects the beginning of the connector from the shape.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3) 
    If .Connector Then 
        With .ConnectorFormat 
            If .BeginConnected Then 
                oldBeginConnSite = .BeginConnectionSite 
                Set oldBeginConnShape = .BeginConnectedShape 
                .BeginDisconnect 
            End If 
        End With 
    End If 
End With
```


## See also


#### Concepts


[ConnectorFormat Object](connectorformat-object-excel.md)

