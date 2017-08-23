---
title: "Метод Application.PointsToPicas (издатель)"
keywords: vbapb10.chm131160
f1_keywords: vbapb10.chm131160
ms.prod: publisher
api_name: Publisher.Application.PointsToPicas
ms.assetid: ff566bef-7032-70f7-7880-ff66cfeca88f
ms.date: 06/08/2017
ms.openlocfilehash: b1417b40c6546fe699a6696f656799c9ccabec16
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationpointstopicas-method-publisher"></a>Метод Application.PointsToPicas (издатель)

Преобразует измерения из точки пики (1 пика = 12 точек). Возвращает преобразованные измерения как **один**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PointsToPicas** ( **_Значение_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **Один**|Значение точки для преобразования в пики.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Используйте метод **[PicasToPoints](application-picastopoints-method-publisher.md)** для преобразования измерений в пики в пунктах.


## <a name="example"></a>Пример

В этом примере выполняется преобразование измерения в пунктах, введенный пользователем измерений в пики.


```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in points (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) &; " points = " _ 
 &; Format(Application _ 
 .PointsToPicas(Value:=Val(strInput)), _ 
 "0.00") &; " picas" 
 
 MsgBox strOutput 
Loop
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

