---
title: "Метод ConnectorFormat.EndConnect (издатель)"
keywords: vbapb10.chm3211282
f1_keywords: vbapb10.chm3211282
ms.prod: publisher
api_name: Publisher.ConnectorFormat.EndConnect
ms.assetid: d37c1ab2-d73a-903b-7c5d-f38a29544728
ms.date: 06/08/2017
ms.openlocfilehash: 8329d0a4fc717e00d9808b0c3bd501c1e2625feb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="connectorformatendconnect-method-publisher"></a>Метод ConnectorFormat.EndConnect (издатель)

Подключает в конец указанный соединитель указанного фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EndConnect** ( **_ConnectedShape_**, **_ConnectionSite_**)

 переменная _expression_A, представляет собой объект- **ConnectorFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|ConnectedShape|Обязательное свойство.| **Фигура**|Фигура, к которому Microsoft Publisher подключает окончании соединителя. Указанный объект **фигуры** должен быть в одном семействе **фигур** , как соединитель.|
|ConnectionSite|Обязательное свойство.| **Длинный**|Подключение сайта на фигуры, указанного идентификатором ConnectedShape. Должно быть целое число от 1 до целое число, возвращаемое свойством **[ConnectionSiteCount](shape-connectionsitecount-property-publisher.md)** указанного фигуры. Подключение сайтов нумеруются, начиная с первого указанного фигуры и против часовой стрелки перемещения фигуры. Если требуется соединитель автоматически найти короткий путь между двумя фигурами, к которому подключен задания любое допустимое целое значение для этого аргумента, а затем используйте метод **[RerouteConnections](shape-rerouteconnections-method-publisher.md)** после присоединения к фигурам с обоих концов соединитель.|

## <a name="remarks"></a>Заметки

Если уже соединение между окончании соединителя и другую фигуру, это подключение будет отключена. Конец соединитель уже не находится в указанной связи сайтов, этот метод перемещает конца соединительной линии связи сайтов и изменяет размер и положение соединителя.

При присоединении соединитель на объект, размер и положение соединителя автоматически настраиваются при необходимости.

Используйте метод **[BeginConnect](connectorformat-beginconnect-method-publisher.md)** присоединение начала соединитель фигуры.


## <a name="example"></a>Пример

В этом примере добавляется два прямоугольника для первой страницы в активной публикации и связывает их с искривленной формы. Обратите внимание на то, что метод **RerouteConnections** переопределяет значения, которые вы задаете **_ConnectionSite_** аргументов, используемых с методами **BeginConnect** и **EndConnect** .


```vb
Dim shpRect1 As Shape 
Dim shpRect2 As Shape 
 
With ActiveDocument.Pages(1).Shapes 
 
 ' Add two new rectangles. 
 Set shpRect1 = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=100, Top:=50, Width:=200, Height:=100) 
 Set shpRect2 = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=300, Top:=300, Width:=200, Height:=100) 
 
 ' Add a new curved connector. 
 With .AddConnector(Type:=msoConnectorCurve, _ 
 BeginX:=0, BeginY:=0, EndX:=100, EndY:=100) _ 
 .ConnectorFormat 
 
 ' Connect the new connector to the two rectangles. 
 .BeginConnect ConnectedShape:=shpRect1, ConnectionSite:=1 
 .EndConnect ConnectedShape:=shpRect2, ConnectionSite:=1 
 
 ' Reroute the connector to create the shortest path. 
 .Parent.RerouteConnections 
 End With 
 
End With
```


