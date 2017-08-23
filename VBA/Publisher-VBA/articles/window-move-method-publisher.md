---
title: "Метод Window.Move (издатель)"
keywords: vbapb10.chm262163
f1_keywords: vbapb10.chm262163
ms.prod: publisher
api_name: Publisher.Window.Move
ms.assetid: a33b213b-6549-abf7-0217-041b469b798a
ms.date: 06/08/2017
ms.openlocfilehash: 52009b8e36cd10fcecc5776420b436583943ab4a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="windowmove-method-publisher"></a>Метод Window.Move (издатель)

Перемещает окно активного документа.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Перемещение** ( **_Слева_**, **_сверху_**)

 переменная _expression_A, представляющий объект **Window** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Слева|Обязательное свойство.| **Длинный**|Экран горизонтальную позицию окна, указанного.|
|Вверх|Обязательное свойство.| **Длинный**|Вертикальная экранная позицию окна, указанного.|

## <a name="remarks"></a>Заметки

Если окно приложения свернуто или развернуто, этот метод возвращает ошибку.


## <a name="example"></a>Пример

В этом примере проверяется состояние окна приложения и если он не развернуто и не свернуто, перемещает окно в верхнем левом углу экрана.


```vb
Sub MoveWindow() 
 With ActiveWindow 
 If .WindowState = pbWindowStateNormal Then 
 .Move Left:=50, Top:=50 
 End If 
 End With 
End Sub
```


