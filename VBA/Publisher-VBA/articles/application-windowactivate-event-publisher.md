---
title: "Событие Application.WindowActivate (издатель)"
keywords: vbapb10.chm268435457
f1_keywords: vbapb10.chm268435457
ms.prod: publisher
api_name: Publisher.Application.WindowActivate
ms.assetid: a7e4e396-9661-763c-8e41-dc279757af94
ms.date: 06/08/2017
ms.openlocfilehash: 301eec8e38279c3db58592e66e2169cf9c9170c9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationwindowactivate-event-publisher"></a>Событие Application.WindowActivate (издатель)

Происходит при активации окна приложения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WindowActivate** ( **_Низ_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Низ|Обязательное свойство.| **Окно**|Окно, вызывается.|

## <a name="remarks"></a>Заметки

Сведения об использовании событий с помощью объекта приложения [С помощью событий объекта](using-events-with-the-application-object-publisher.md)см.


## <a name="example"></a>Пример

В этом примере разворачивает окно Microsoft Publisher при его активации. Этот код должен находиться в модуле класса и экземпляр класса необходимо правильно инициализировать данный пример работы; просмотра [Событий с помощью объекта](using-events-with-the-application-object-publisher.md)для указания о том, как это сделать.


```vb
Public WithEvents appPublisher as Publisher.Application 
 
Private Sub appPublisher_WindowActivate _ 
 (ByVal Wn As Window) 
 Wn.WindowState = pbWindowStateMaximize 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

