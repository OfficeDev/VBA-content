---
title: "Метод Shapes.AddCurve (издатель)"
keywords: vbapb10.chm2162706
f1_keywords: vbapb10.chm2162706
ms.prod: publisher
api_name: Publisher.Shapes.AddCurve
ms.assetid: 888a35cb-190d-4058-e0d7-a848d77ba920
ms.date: 06/08/2017
ms.openlocfilehash: 43031eae70a33cf3f9feeb54e02b6b74aa0c82a4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesaddcurve-method-publisher"></a>Метод Shapes.AddCurve (издатель)

Добавляет новый объект **[фигуры](shape-object-publisher.md)** , предоставляющий Безье график для указанной коллекции **[фигур](shapes-object-publisher.md)** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddCurve** ( **_SafeArrayOfPoints_**)

 переменная _expression_A, представляет собой объект- **фигур** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|SafeArrayOfPoints|Обязательное свойство.| **Variant**|Массив координации пары, указывает грани и контроль точек график. Первой точкой указываемые является отправной вершин и следующие две точки являются контрольными для первого сегмента Безье. Для каждого дополнительного сегмента график укажите узел и двух контрольных точек. Конечная точка, указываемые — это последний вершины график. Обратите внимание на то, что всегда следует указывать 3n + 1 точки, где n — число сегментов в поверхности.|

### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="remarks"></a>Заметки

Для элементов массива в **_SafeArrayOfPoints_**числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).


## <a name="example"></a>Пример

Следующий пример добавляет два сегмента Безье график для первой страницы публикации active.


```vb
Dim shpCurve As Shape 
Dim arrPoints(1 To 4, 1 To 2) As Single 
 
arrPoints(1, 1) = 0 
arrPoints(1, 2) = 0 
arrPoints(2, 1) = 72 
arrPoints(2, 2) = 72 
arrPoints(3, 1) = 144 
arrPoints(3, 2) = 36 
arrPoints(4, 1) = 216 
arrPoints(4, 2) = 108 
 
Set shpCurve = ActiveDocument.Pages(1).Shapes.AddCurve _ 
 (SafeArrayOfPoints:=arrPoints)
```


