---
title: "Метод Application.PicasToPoints (издатель)"
keywords: vbapb10.chm131152
f1_keywords: vbapb10.chm131152
ms.prod: publisher
api_name: Publisher.Application.PicasToPoints
ms.assetid: 64d3e435-dcc1-d637-7aac-cc9a9bf81e76
ms.date: 06/08/2017
ms.openlocfilehash: ad8f76610ec177f019217863081af009b6235a0c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationpicastopoints-method-publisher"></a>Метод Application.PicasToPoints (издатель)

Преобразует измерения из пики в точках (1 пика = 12 точек). Возвращает преобразованные измерения как **один**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PicasToPoints** ( **_Значение_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **Один**|Значение пика для преобразования в пунктах.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Используйте метод **[PointsToPicas](application-pointstopicas-method-publisher.md)** для преобразования значения в пунктах пики.


## <a name="example"></a>Пример

В этом примере выполняется преобразование размеры в пики, введенный пользователем измерений в пунктах.


```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in picas (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) &; " picas = " _ 
 &; Format(Application _ 
 .Picas ToPoints(Value:=Val(strInput)), _ 
 "0.00") &; " points" 
 
 MsgBox strOutput 
Loop
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

