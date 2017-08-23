---
title: "Метод Shapes.AddPolyline (издатель)"
keywords: vbapb10.chm2162711
f1_keywords: vbapb10.chm2162711
ms.prod: publisher
api_name: Publisher.Shapes.AddPolyline
ms.assetid: d49fb2bc-4df5-fff8-c741-2c0d35413fc5
ms.date: 06/08/2017
ms.openlocfilehash: ff794bf4979771bc331460348afe5cf79b49dd5b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesaddpolyline-method-publisher"></a>Метод Shapes.AddPolyline (издатель)

Добавление нового объекта **Shape** , представляющее open ломаной или закрытой многоугольника определенной коллекции **фигур** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddPolyline** ( **_SafeArrayOfPoints_**)

 переменная _expression_A, представляет собой объект- **фигур** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|SafeArrayOfPoints|Обязательное свойство.| **Variant**|Массив пар координат, указывающее вершины ломаной пользователя или многоугольника.|

### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="remarks"></a>Заметки

Для элементов массива в **_SafeArrayOfPoints_**числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).

Чтобы закрытой многоугольник, назначьте те же координаты грани имени и фамилии в ломаной документа.


## <a name="example"></a>Пример

В следующем примере добавляется треугольник для первой страницы active публикации. Поскольку начальную и конечную точки имеют те же координаты, многоугольника закрывается.


```vb
Dim shpPolyline As Shape 
Dim arrPoints(1 To 4, 1 To 2) As Single 
 
arrPoints(1, 1) = 25 
arrPoints(1, 2) = 100 
arrPoints(2, 1) = 100 
arrPoints(2, 2) = 150 
arrPoints(3, 1) = 150 
arrPoints(3, 2) = 50 
arrPoints(4, 1) = 25 
arrPoints(4, 2) = 100 
 
Set shpPolyline = ActiveDocument.Pages(1).Shapes.AddPolyline _ 
 (SafeArrayOfPoints:=arrPoints)
```


