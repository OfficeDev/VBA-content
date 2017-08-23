---
title: "Метод ConnectorFormat.EndDisconnect (издатель)"
keywords: vbapb10.chm3211283
f1_keywords: vbapb10.chm3211283
ms.prod: publisher
api_name: Publisher.ConnectorFormat.EndDisconnect
ms.assetid: cf46820e-99fc-cdd2-999a-f30117b5616f
ms.date: 06/08/2017
ms.openlocfilehash: 311f3b31562e3515456bd989fc275eff130b2150
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="connectorformatenddisconnect-method-publisher"></a>Метод ConnectorFormat.EndDisconnect (издатель)

Отключает в конец указанный соединитель из формы, к которой он подключен.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EndDisconnect**

 переменная _expression_A, представляет собой объект- **ConnectorFormat** .


## <a name="remarks"></a>Заметки

Этот метод не изменяет размер и положение соединителя; конец соединитель остается расположенных на сайте подключения, но больше не подключен.

Используйте метод **[BeginDisconnect](connectorformat-begindisconnect-method-publisher.md)** для отключения начала соединителя из фигуры.


## <a name="example"></a>Пример

В этом примере добавляет два прямоугольника в первой страницы в активной публикации, связывает их с соединитель, автоматически перенаправляет соединителя Минимальная пути и затем отключает соединитель из прямоугольники.


```vb
Dim shpRect1 As Shape 
Dim shpRect2 As Shape 
 
With ActiveDocument.Pages(1).Shapes 
 
 ' Add two new rectangles. 
 Set shpRect1 = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=100, Top:=50, Width:=200, Height:=100) 
 Set shpRect2 = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=300, Top:=300, Width:=200, Height:=100) 
 
 ' Add a new connector. 
 With .AddConnector(Type:=msoConnectorCurve, _ 
 BeginX:=0, BeginY:=0, EndX:=0, EndY:=0) _ 
 .ConnectorFormat 
 
 ' Connect the new connector to the two rectangles. 
 .BeginConnect ConnectedShape:=shpRect1, ConnectionSite:=1 
 .EndConnect ConnectedShape:=shpRect2, ConnectionSite:=1 
 
 ' Reroute the connector to create the shortest path. 
 .Parent.RerouteConnections 
 
 ' Disconnect the new connector from the rectangles but 
 ' leave in place. 
 .BeginDisconnect 
 .EndDisconnect 
 End With 
 
End With
```


