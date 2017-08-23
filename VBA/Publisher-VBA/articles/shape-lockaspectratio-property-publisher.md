---
title: "Свойство Shape.LockAspectRatio (издатель)"
keywords: vbapb10.chm2228291
f1_keywords: vbapb10.chm2228291
ms.prod: publisher
api_name: Publisher.Shape.LockAspectRatio
ms.assetid: eeb87bb5-01d5-5d21-b268-045497ea3682
ms.date: 06/08/2017
ms.openlocfilehash: a542b2a455e081ca4bfbc6ed83f1e4291ba9abd5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapelockaspectratio-property-publisher"></a>Свойство Shape.LockAspectRatio (издатель)

Возвращает или задает константой **MsoTriState**, указывающее, является ли указанный фигуры сохраняются исходные пропорции при изменении размера. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **LockAspectRatio**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="remarks"></a>Заметки

Значение свойства **LockAspectRatio** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Высота и ширина формы изменения независимо друг от друга при изменении размера.|
| **msoTriStateMixed**|Возвращает значение, указывающее, сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**|Указанный фигуры сохраняются исходные пропорции при изменении размера.|

## <a name="example"></a>Пример

В этом примере добавляется куба active публикацию. Куб можно переместить и размера, но не reproportioned.


```vb
Dim shp As Shape 
 
Set shp = ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeCube, _ 
 Left:=50, Top:=50, Width:=100, Height:=200) _ 
 
shp.LockAspectRatio = msoTrue
```


