---
title: "Свойство Shape.WebOptionButton (издатель)"
keywords: vbapb10.chm2228343
f1_keywords: vbapb10.chm2228343
ms.prod: publisher
api_name: Publisher.Shape.WebOptionButton
ms.assetid: 0c43387c-0cb6-5d6f-68cb-d1883ce17243
ms.date: 06/08/2017
ms.openlocfilehash: 86a86e1021281f84924759d3eb8bfb0cba96e898
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeweboptionbutton-property-publisher"></a>Свойство Shape.WebOptionButton (издатель)

Возвращает объект **[WebOptionButton](weboptionbutton-object-publisher.md)** , связанный с указанным фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WebOptionButton**

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="return-value"></a>Возвращаемое значение

WebOptionButton


## <a name="example"></a>Пример

В этом примере создается новая кнопка параметр Web и указывает, что выбран пункт состояние по умолчанию.


```vb
Dim shpNew As Shape 
Dim wobTemp As WebOptionButton 
 
Set shpNew = ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlOptionButton, Left:=100, _ 
 Top:=123, Width:=16, Height:=10) 
 
Set wobTemp = shpNew.WebOptionButton 
 
wobTemp.Selected = msoTrue
```


