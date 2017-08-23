---
title: "Свойство Shape.ConnectionSiteCount (издатель)"
keywords: vbapb10.chm2228276
f1_keywords: vbapb10.chm2228276
ms.prod: publisher
api_name: Publisher.Shape.ConnectionSiteCount
ms.assetid: 00c32910-96b6-6981-8359-de4a71852934
ms.date: 06/08/2017
ms.openlocfilehash: 35c0b0136de9007db65782b4a67bd08695680132
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeconnectionsitecount-property-publisher"></a>Свойство Shape.ConnectionSiteCount (издатель)

Возвращает значение типа **Long** , показывающее общее число подключений к сайтам на текущий объект **фигуры** . Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ConnectionSiteCount**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="remarks"></a>Заметки

Число сайтов подключения зависит от того, Геометрия фигуры. Прямоугольный объекты, включая таблицы и веб-элементы управления вероятнее всего будут иметь четыре узлами, один по центру на каждом пограничном фигуры.


## <a name="example"></a>Пример

В этом примере добавляется два прямоугольника active публикации и соединяет их с двумя разъемами. Основные компоненты оба соединители с подключением к сайту подключения одно на первый прямоугольник; заканчивается соединители с подключением к первого и последнего подключения к сайтам второго прямоугольника. Затем подсчитывает число подключений на первый прямоугольник.


```vb
Sub Connections() 
 
 Dim shpNew As Shapes 
 Dim shpFirstRect As Shape 
 Dim shpSecondRect As Shape 
 Dim intLastSite As Integer 
 Dim intCount As Integer 
 
 Set shpNew = Application.ActiveDocument _ 
 .MasterPages(Item:=1).Shapes 
 Set shpFirstRect = shpNew.AddShape(Type:=msoShapeRectangle, _ 
 Left:=100, Top:=50, Width:=200, Height:=100) 
 Set shpSecondRect = shpNew.AddShape(msoShapeRectangle, _ 
 Left:=300, Top:=300, Width:=200, Height:=100) 
 varLastSite = shpSecondRect.ConnectionSiteCount 
 
 ' Add the first connector from rectangle 1, 
 ' site 1 to rectangle 2, site 1. 
 With shpNew.AddConnector(Type:=msoConnectorCurve, _ 
 BeginX:=0, BeginY:=0, EndX:=100, EndY:=100) _ 
 .ConnectorFormat 
 .BeginConnect ConnectedShape:=shpFirstRect, ConnectionSite:=1 
 .EndConnect ConnectedShape:=shpSecondRect, ConnectionSite:=1 
 End With 
 
 ' Add the second connector from rectangle 1, 
 ' site 1 to rectangle 2, site 2. 
 With shpNew.AddConnector(Type:=msoConnectorCurve, _ 
 BeginX:=0, BeginY:=0, EndX:=100, EndY:=100) _ 
 .ConnectorFormat 
 .BeginConnect ConnectedShape:=shpFirstRect, ConnectionSite:=1 
 .EndConnect ConnectedShape:=shpSecondRect, _ 
 ConnectionSite:=intLastSite 
 End With 
 
 intCount = shpFirstRect.ConnectionSiteCount 
 
End Sub
```


