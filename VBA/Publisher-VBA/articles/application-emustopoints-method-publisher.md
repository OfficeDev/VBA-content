---
title: "Метод Application.EmusToPoints (издатель)"
keywords: vbapb10.chm131142
f1_keywords: vbapb10.chm131142
ms.prod: publisher
api_name: Publisher.Application.EmusToPoints
ms.assetid: 941e5975-ca7a-38dc-8116-e90b2a2ab6e5
ms.date: 06/08/2017
ms.openlocfilehash: ac2533aa3c0da84518914cdd3e690fd8a0c8fabc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationemustopoints-method-publisher"></a>Метод Application.EmusToPoints (издатель)

Преобразует измерения из emus точек (12700 emus = 1 пункт). Возвращает преобразованные измерения как **один**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EmusToPoints** ( **_Значение_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Значение|Обязательное свойство.| **Один**|Выражение, возвращающее один из объектов в списке применяется к.|

### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Используйте метод **[PointsToEmus](application-pointstoemus-method-publisher.md)** для преобразования emus измерения в точках.


## <a name="example"></a>Пример

В этом примере выполняется преобразование размеры в emus, введенный пользователем измерений в пунктах.


```vb
Dim strInput As String 
Dim strOutput As String 
 
Do While True 
 ' Get input from user. 
 strInput = InputBox( _ 
 Prompt:="Enter measurement in emus (0 to cancel): ", _ 
 Default:="0") 
 
 ' Exit the loop if user enters zero. 
 If Val(strInput) = 0 Then Exit Do 
 
 ' Evaluate and display result. 
 strOutput = Trim(strInput) &; " emus = " _ 
 &; Format(Application _ 
 .EmusToPoints(Value:=Val(strInput)), _ 
 "0.00") &; " points" 
 
 MsgBox strOutput 
Loop 

```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

