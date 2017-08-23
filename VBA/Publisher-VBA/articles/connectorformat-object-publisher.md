---
title: "Объект ConnectorFormat (издатель)"
keywords: vbapb10.chm3276799
f1_keywords: vbapb10.chm3276799
ms.prod: publisher
api_name: Publisher.ConnectorFormat
ms.assetid: 9b541d54-b1b9-c023-c9c4-08ff6b811eb9
ms.date: 06/08/2017
ms.openlocfilehash: 4a599259c84cc1b08b9608577dce69d6b8928034
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="connectorformat-object-publisher"></a>Объект ConnectorFormat (издатель)

Содержит свойства и методы, которые применяются к соединители. Соединитель — это строка, которая связывает две фигуры при вызове узлами уровнях. При реорганизации фигуры, которые подключаются геометрии соединитель будет автоматически настроен таким образом, фигур, остаются соединенными.
 


## <a name="example"></a>Пример

Используйте свойство **ConnectorFormat** объекта **[Shape](shape-object-publisher.md)** или коллекции **[ShapeRange](shaperange-object-publisher.md)** возвращает объект **ConnectorFormat** . Используйте методы **[BeginConnect](connectorformat-beginconnect-method-publisher.md)** и **[EndConnect](connectorformat-endconnect-method-publisher.md)** объекта **ConnectorFormat** присоединение завершается соединитель на другие фигуры в публикации. Используйте метод **[RerouteConnections](shape-rerouteconnections-method-publisher.md)** объекта **Shape** и семейство сайтов **ShapeRange** автоматически найти короткий путь между двумя фигурами, подключенных по соединитель. Используйте свойство **[соединителя](shape-connector-property-publisher.md)** ли соединитель фигуры.
 

 

 

 
Обратите внимание на то, назначьте размер и положение при добавлении соединитель в коллекции **фигур** , но размер и положение автоматически настраиваются при присоединении начала и окончания соединителя на другие фигуры в коллекции. Таким образом Если вы намереваетесь соединитель с подключением к другим фигурам, исходный размер и положение указываемые не имеют значения. Аналогичным образом укажите какие сайты подключения на форму на Подключите разъем при подключении соединитель, но с помощью метода **RerouteConnections** после присоединения соединителя может изменить подключение сайтов соединитель подключает, что исходный вариант подключения сайтов имеют значения.
 

 

 

 
В следующем примере добавляется два прямоугольника active публикации и связывает их с искривленной формы.
 

 



```
Dim shpAll As Shapes 
Dim firstRect As Shape 
Dim secondRect As Shape 
 
Set shpAll = ActiveDocument.Pages(1).Shapes 
Set firstRect = shpAll.AddShape(Type:=msoShapeRectangle, _ 
 Left:=100, Top:=50, Width:=200, Height:=100) 
Set secondRect = shpAll.AddShape(Type:=msoShapeRectangle, _ 
 Left:=300, Top:=300, Width:=200, Height:=100) 

```




```
With shpAll.AddConnector(Type:=msoConnectorCurve, BeginX:=0, _ 
 BeginY:=0, EndX:=0, EndY:=0).ConnectorFormat 
 .BeginConnect ConnectedShape:=firstRect, ConnectionSite:=1 
 .EndConnect ConnectedShape:=secondRect, ConnectionSite:=1 
 .Parent.RerouteConnections 
End With
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[BeginConnect](connectorformat-beginconnect-method-publisher.md)|
|[BeginDisconnect](connectorformat-begindisconnect-method-publisher.md)|
|[EndConnect](connectorformat-endconnect-method-publisher.md)|
|[EndDisconnect](connectorformat-enddisconnect-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](connectorformat-application-property-publisher.md)|
|[BeginConnected](connectorformat-beginconnected-property-publisher.md)|
|[BeginConnectedShape](connectorformat-beginconnectedshape-property-publisher.md)|
|[BeginConnectionSite](connectorformat-beginconnectionsite-property-publisher.md)|
|[EndConnected](connectorformat-endconnected-property-publisher.md)|
|[EndConnectedShape](connectorformat-endconnectedshape-property-publisher.md)|
|[EndConnectionSite](connectorformat-endconnectionsite-property-publisher.md)|
|[Родительский раздел](connectorformat-parent-property-publisher.md)|
|[Type](connectorformat-type-property-publisher.md)|

