---
title: "Свойство Options.HyphenationZone (издатель)"
keywords: vbapb10.chm1048593
f1_keywords: vbapb10.chm1048593
ms.prod: publisher
api_name: Publisher.Options.HyphenationZone
ms.assetid: ed0e90de-4a2a-3c8a-27f1-e8c7c1f0e174
ms.date: 06/08/2017
ms.openlocfilehash: c7c71597ca57546043d0d648828ead08a12b6f89
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionshyphenationzone-property-publisher"></a>Свойство Options.HyphenationZone (издатель)

Возвращает или задает **Variant** , который представляет максимальный объем пространства, который Microsoft Publisher оставляет между окончания последнего слова в строке и правого поля. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HyphenationZone**

 переменная _expression_A, представляет собой объект- **Параметры** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="example"></a>Пример

В этом примере показано включение автоматической расстановки переносов и указывает максимальный объем пространства между окончания последнего слова и правого поля, равное 1 дюйм (72 точки).


```vb
Sub SetHyphenationZone() 
 With Options 
 .AutoHyphenate = True 
 .HyphenationZone = 72 
 End With 
End Sub
```


