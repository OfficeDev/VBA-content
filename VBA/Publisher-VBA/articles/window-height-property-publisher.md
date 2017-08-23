---
title: "Свойство Window.Height (издатель)"
keywords: vbapb10.chm262151
f1_keywords: vbapb10.chm262151
ms.prod: publisher
api_name: Publisher.Window.Height
ms.assetid: 3d47bb99-bab7-b5aa-c834-04bcd6e8b151
ms.date: 06/08/2017
ms.openlocfilehash: a5a7beb05adcfd93c3321b4b58cf7b3656cbeb99
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="windowheight-property-publisher"></a>Свойство Window.Height (издатель)

Возвращает или задает **Long** , представляющее высоту окна (в точках). Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Высота**

 переменная _expression_A, представляющий объект **Window** .


## <a name="remarks"></a>Заметки

Допустимые значения для свойства **Height** зависит от размера рабочей области приложения и позиции объекта в рабочей области. По центру объектов на размер страницы не баннер свойство **Height** может быть 0,0-50,0 дюйма. По центру объектов на размер заголовка страницы свойство **Height** может быть 0.0 для 241.0 дюйма.


## <a name="example"></a>Пример

Этот пример устанавливает высоту и ширину окна, если окно не развернуто и не свернуто.


```vb
Sub SetWindowHeight() 
 With ActiveWindow 
 If .WindowState <> pbWindowStateNormal Then 
 .WindowState = pbWindowStateNormal 
 .Height = InchesToPoints(5) 
 .Width = InchesToPoints(5) 
 End If 
 End With 
End Sub
```


