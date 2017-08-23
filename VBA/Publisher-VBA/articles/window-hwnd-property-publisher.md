---
title: "Свойство Window.Hwnd (издатель)"
keywords: vbapb10.chm262161
f1_keywords: vbapb10.chm262161
ms.prod: publisher
api_name: Publisher.Window.Hwnd
ms.assetid: e0fe9b33-0839-a2a5-f939-9906e46f9632
ms.date: 06/08/2017
ms.openlocfilehash: c5f1b64e47a933f649592308a62ab967d3fc758c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="windowhwnd-property-publisher"></a>Свойство Window.Hwnd (издатель)

Возвращает значение типа **Long** , указывающее, дескриптор окна приложения Microsoft Publisher. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HWND**

 переменная _expression_A, представляющий объект **Window** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="example"></a>Пример

Следующий пример отображает дескриптор окна приложения Publisher.


```vb
MsgBox "The handle to the Publisher application window is " &; _ 
 Application.ActiveWindow.Hwnd
```


