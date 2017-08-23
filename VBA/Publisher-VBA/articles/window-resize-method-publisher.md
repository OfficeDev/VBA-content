---
title: "Метод Window.Resize (издатель)"
keywords: vbapb10.chm262164
f1_keywords: vbapb10.chm262164
ms.prod: publisher
api_name: Publisher.Window.Resize
ms.assetid: 478e5f05-a2f9-c3b0-5dd0-3248272b2c37
ms.date: 06/08/2017
ms.openlocfilehash: 21d7dd6f40e6da63c1d067887450670de9561251
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="windowresize-method-publisher"></a>Метод Window.Resize (издатель)

Изменение размеров окна приложения Microsoft Publisher.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Чтобы изменить размер** ( **_Ширина_**, **_Высота_**)

 переменная _expression_A, представляющий объект **Window** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Width|Обязательное свойство.| **Длинный**|Ширина окна в точках.|
|Height|Обязательное свойство.| **Длинный**|Высота окна в точках.|

## <a name="remarks"></a>Заметки

Если окно свернуто или развернуто, возникает ошибка.

Использование свойств **[ширины](window-width-property-publisher.md)** и **[высоты](window-height-property-publisher.md)** Установка высоты и ширины окна независимо друг от друга.


## <a name="example"></a>Пример

В этом примере изменяет размер окна приложения Publisher 7 дюймов широкий с высокой 6 дюймов.


```vb
With Application.ActiveWindow 
 .WindowState = wdWindowStateNormal 
 .Resize Width:=InchesToPoints(7), Height:=InchesToPoints(6) 
End With
```


