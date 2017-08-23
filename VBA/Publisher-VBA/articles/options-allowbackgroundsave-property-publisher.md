---
title: "Свойство Options.AllowBackgroundSave (издатель)"
keywords: vbapb10.chm1048577
f1_keywords: vbapb10.chm1048577
ms.prod: publisher
api_name: Publisher.Options.AllowBackgroundSave
ms.assetid: 5bddfb2d-7fb7-99db-43ea-c6ee53e1d0b3
ms.date: 06/08/2017
ms.openlocfilehash: a439aab747b1950ae4177ae017c54bb616a60e23
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionsallowbackgroundsave-property-publisher"></a>Свойство Options.AllowBackgroundSave (издатель)

 **Значение true** (по умолчанию) для Microsoft Publisher для сохранения публикации в фоновом режиме, позволяя пользователям выполнять другие действия в то же время. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AllowBackgroundSave**

 переменная _expression_A, представляющий объект **параметров** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Этот параметр сохраняется для каждого пользователя и сохраняет из одного сеанса.


## <a name="example"></a>Пример

В этом примере показано отключение фон сохранить, поэтому публикаций не следует сохранять в фоновом режиме.


```vb
Sub DoNotSaveInBackground() 
 Options.AllowBackgroundSave = False 
End Sub
```


