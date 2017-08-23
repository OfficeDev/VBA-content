---
title: "Метод Options.ResetTips (издатель)"
keywords: vbapb10.chm1048616
f1_keywords: vbapb10.chm1048616
ms.prod: publisher
api_name: Publisher.Options.ResetTips
ms.assetid: a119aacc-ba19-f430-e8af-6d84c438ec25
ms.date: 06/08/2017
ms.openlocfilehash: c88f434c76448cd1273023e581912e7ac3d634ce
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionsresettips-method-publisher"></a>Метод Options.ResetTips (издатель)

Сброс страницы советов, чтобы пользователь может просматривать их при использовании функции, которые использовались перед.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ResetTips**

 переменная _expression_A, представляющий объект **параметров** .


## <a name="remarks"></a>Заметки

Метод **ResetTips** эквивалентно нажав кнопку **Сброс советы** на вкладке **Помощь пользователю** диалоговое окно " **Параметры** " (меню " **Сервис** ").


## <a name="example"></a>Пример

В этом примере восстанавливаются значения по умолчанию всплывающие подсказки.


```vb
Sub ResetTippages() 
 Options.ResetTips 
End Sub
```


