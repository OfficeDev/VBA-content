---
title: "Метод Shape.RerouteConnections (издатель)"
keywords: vbapb10.chm2228260
f1_keywords: vbapb10.chm2228260
ms.prod: publisher
api_name: Publisher.Shape.RerouteConnections
ms.assetid: 04afd4aa-dc84-d39c-e9fa-d06f8f4c0a02
ms.date: 06/08/2017
ms.openlocfilehash: fdbad633b827119cd07b751c8a3f9bb0ba4e0fd6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapererouteconnections-method-publisher"></a>Метод Shape.RerouteConnections (издатель)

Изменение пути соединители, чтобы они вступили Минимальная возможные пути между фигурами, которые они подключаются. Для этого метода **RerouteConnections** может отсоединить концах соединитель и присоедините их различных связи сайтов на присоединенными фигурами.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RerouteConnections**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="remarks"></a>Заметки

Этот метод перенаправляет все соединители, подключенного к указанной фигуры; Если указанный фигуры соединитель, пересылаются его.


## <a name="example"></a>Пример

В этом примере добавляется два прямоугольника для первой страницы в активной публикации и связывает их с искривленной формы. Обратите внимание на то, что метод **RerouteConnections** переопределяет значения, которые вы задаете **_ConnectionSite_** аргументов, используемых с методами **BeginConnect**и **EndConnect** .


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


