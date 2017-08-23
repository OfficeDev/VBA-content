---
title: "Свойство ConnectorFormat.BeginConnectedShape (издатель)"
keywords: vbapb10.chm3211521
f1_keywords: vbapb10.chm3211521
ms.prod: publisher
api_name: Publisher.ConnectorFormat.BeginConnectedShape
ms.assetid: a7eb9090-ad01-234c-99ff-3bb0616d02c0
ms.date: 06/08/2017
ms.openlocfilehash: 75bdd5284901d84e45b72f43c1fc8c4fcbd28e6a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="connectorformatbeginconnectedshape-property-publisher"></a>Свойство ConnectorFormat.BeginConnectedShape (издатель)

Возвращает объект **[фигуры](shape-object-publisher.md)** , представляющий фигуры, к которому подключен начала указанный соединитель.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BeginConnectedShape**

 переменная _expression_A, представляет собой объект- **ConnectorFormat** .


### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="remarks"></a>Заметки

Если в начало указанный соединитель не будет присоединен к фигуры, возникает ошибка.

Свойство **[EndConnectedShape](connectorformat-endconnectedshape-property-publisher.md)** возвращает фигуры, подключенного к конца соединителя.


## <a name="example"></a>Пример

В этом примере предполагается, что первая страница в активной публикации уже содержит две фигуры, подключенное соединителем с именем Conn1To2. Код добавляет прямоугольник и соединитель для первой страницы. В начало новый соединитель будет присоединена к на одном узле подключения в начале соединитель с именем Conn1To2 и end новый соединитель будет присоединена к точке, один на новый прямоугольник.


```vb
Dim shpNew As Shape 
Dim intSite As Integer 
Dim shpOld As Shape 
 
With ActiveDocument.Pages(1).Shapes 
 
 ' Add new rectangle. 
 Set shpNew = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=450, Top:=190, Width:=200, Height:=100) 
 
 ' Add new connector. 
 .AddConnector(Type:=msoConnectorCurve, _ 
 BeginX:=0, BeginY:=0, EndX:=10, EndY:=10) _ 
 .Name = "Conn1To3" 
 
 ' Get connection site number of old shape, and set 
 ' reference to old shape. 
 With .Item("Conn1To2").ConnectorFormat 
 intSite = .BeginConnectionSite 
 Set shpOld = .BeginConnectedShape 
 End With 
 
 ' Connect new connector to old shape and new rectangle. 
 With .Item("Conn1To3").ConnectorFormat 
 .BeginConnect ConnectedShape:=shpOld, _ 
 ConnectionSite:=intSite 
 .EndConnect ConnectedShape:=shpNew, _ 
 ConnectionSite:=1 
 End With 
End With 

```


