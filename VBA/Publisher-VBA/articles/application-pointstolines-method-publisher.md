---
title: "Метод Application.PointsToLines (издатель)"
keywords: vbapb10.chm131158
f1_keywords: vbapb10.chm131158
ms.prod: publisher
api_name: Publisher.Application.PointsToLines
ms.assetid: beab39fe-9458-6878-ae45-487a8b2271df
ms.date: 06/08/2017
ms.openlocfilehash: 64ae9836422a6131d801b8cf290c087224365bae
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationpointstolines-method-publisher"></a>Метод Application.PointsToLines (издатель)

Преобразует измерения из точки на линии (1 строка = 12 точек). Возвращает преобразованные измерения как **один**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PointsToLines** ( **_Значение_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **Один**|Значение точки для преобразования в строки.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Этот метод предполагает измерений в строках 12 пунктов, фактический размер любого текста в публикации не оказывает влияния на коэффициент преобразования.

Используйте метод **[LinesToPoints](application-linestopoints-method-publisher.md)** для преобразования измерений в строках в пунктах.


## <a name="example"></a>Пример

В этом примере преобразует измерений в строках измерения в пунктах, демонстрирующие не влияет на коэффициент преобразования на наличие размер шрифта в текущем выборе. В активной публикации для работы этого примера необходимо выбрать какой-либо текст.


```vb
Dim strOutput As String 
 
' Set text size to 10 points. 
Selection.TextRange.Font.Size = 10 
 
' Display result for 12 points. 
strOutput = "12 points = " _ 
 &; Format(Application _ 
 .PointsToLines(Value:=12), _ 
 "0.00") &; " lines"
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

