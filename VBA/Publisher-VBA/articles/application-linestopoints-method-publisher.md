---
title: "Метод Application.LinesToPoints (издатель)"
keywords: vbapb10.chm131144
f1_keywords: vbapb10.chm131144
ms.prod: publisher
api_name: Publisher.Application.LinesToPoints
ms.assetid: 55c531aa-5619-6f7f-54e7-7721cb70640e
ms.date: 06/08/2017
ms.openlocfilehash: b70f54211926f65e9038b1037b96e884f17f47ae
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationlinestopoints-method-publisher"></a>Метод Application.LinesToPoints (издатель)

Преобразует измерения из строки в пунктах (1 строка = 12 точек). Возвращает преобразованные измерения как **один**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **LinesToPoints** ( **_Значение_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **Один**|Значение строки, который следует преобразовать в пунктах.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Этот метод предполагает измерений в строках 12 пунктов, фактический размер любого текста в публикации не оказывает влияния на коэффициент преобразования.

Используйте метод **[PointsToLines](application-pointstolines-method-publisher.md)** для преобразования строки измерения в точках.


## <a name="example"></a>Пример

В этом примере преобразует измерений в строках измерения в пунктах, демонстрирующие не влияет на коэффициент преобразования на наличие размер шрифта в текущем выборе. В активной публикации для работы этого примера необходимо выбрать какой-либо текст.


```vb
Dim strOutput As String 
 
' Set text size to 10 points. 
Selection.TextRange.Font.Size = 10 
 
' Display result for one line of text. 
strOutput = "1 line = " _ 
 &; Format(Application _ 
 .LinesToPoints(Value:=1), _ 
 "0.00") &; " points"
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

