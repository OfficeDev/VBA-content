---
title: "Свойство Application.ScreenUpdating (издатель)"
keywords: vbapb10.chm131107
f1_keywords: vbapb10.chm131107
ms.prod: publisher
api_name: Publisher.Application.ScreenUpdating
ms.assetid: d265b4fb-1452-91a5-32fe-0cad54c8f29c
ms.date: 06/08/2017
ms.openlocfilehash: 03e533dc1d5dbe286c7be892ee20323e858beec1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationscreenupdating-property-publisher"></a>Свойство Application.ScreenUpdating (издатель)

Возвращает или задает значение **Boolean** , указывающее, является ли Microsoft Publisher обновляет экрана во время выполнения; **Значение true,** чтобы обновить на экране. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Свойство ScreenUpdating**

 переменная _expression_A, представляющий объект **приложения** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Отключение обновления экрана во время выполнения может ускорить выполнение кода Microsoft Visual Basic. Тем не менее рекомендуется предоставить некоторые указания состояния, пользователя принять во внимание, что программа работает правильно.


## <a name="example"></a>Пример

Следующий пример отключает обновление экрана в начале подпрограммы и включает его обратно в конце подпрограммы.


```vb
Sub TurnOffScreenUpdating() 
 ScreenUpdating = False 
 
 ' Execute code here. 
 
 ScreenUpdating = True 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

