---
title: "Метод RulerGuides.Add (издатель)"
keywords: vbapb10.chm720900
f1_keywords: vbapb10.chm720900
ms.prod: publisher
api_name: Publisher.RulerGuides.Add
ms.assetid: 3986452a-73da-04c2-4e11-8369d61cd974
ms.date: 06/08/2017
ms.openlocfilehash: 5f30c860982f3bee9df853ffe75a886c26a6becb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="rulerguidesadd-method-publisher"></a>Метод RulerGuides.Add (издатель)

Добавление направляющей линейки для указанной коллекции **RulerGuides** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Добавление** ( **_Позицию_**, **_Тип_**)

 переменная _expression_A, представляет собой объект- **RulerGuides** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Position|Обязательное свойство.| **Variant**|Положение относительно левого края или верхнего края страницы, в которую будет добавлен новый направляющей линейки. Числовые значения вычисляются в точках; строки вычисляются в единицах, заданных и может быть в любой единицы измерения, поддерживаются в Microsoft Publisher (например, «2,5 дюйма»).|
|Тип|Обязательный| **PbRulerGuideType**|Тип направляющей линейки для добавления.|

## <a name="remarks"></a>Заметки

Тип может иметь одно из следующих констант **PbRulerGuideType** .



| **pbRulerGuideTypeHorizontal**|| **pbRulerGuideTypeVertical**|

## <a name="example"></a>Пример

В следующем примере добавляется направляющие на одну страницу, которые являются 0,5 дюйма от левого верхнего края страницы.


```vb
With ActiveDocument.Pages(1).RulerGuides 
 .Add Position:="0.5 in", Type:=pbRulerGuideTypeHorizontal 
 .Add Position:="0.5 in", Type:=pbRulerGuideTypeVertical 
End With
```


