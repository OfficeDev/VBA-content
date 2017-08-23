---
title: "Метод Application.Quit (издатель)"
keywords: vbapb10.chm131129
f1_keywords: vbapb10.chm131129
ms.prod: publisher
api_name: Publisher.Application.Quit
ms.assetid: db5a02ec-e553-6de1-0e2c-4a9a512e68fe
ms.date: 06/08/2017
ms.openlocfilehash: ddc9c1b9deb86e6a4599fb6951322dbe3ea59c5d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationquit-method-publisher"></a>Метод Application.Quit (издатель)

Выход из программы Microsoft Publisher. Это эквивалентно, нажав кнопку **Выход** в меню **файл** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Завершите работу**

 переменная _expression_A, представляющий объект **приложения** .


## <a name="remarks"></a>Заметки

Чтобы избежать потери несохраненных изменений, используйте метод **[Сохранить](document-save-method-publisher.md)** или **[Сохранить как](document-saveas-method-publisher.md)** для сохранения любого открыть публикацию, прежде чем вызывать метод **Quit** .


## <a name="example"></a>Пример

В этом примере сохраняет открыть публикацию, если он существует и закрывает Publisher.


```vb
If Not (ActiveDocument Is Nothing) 
 ActiveDocument.Save 
End If 
Application.Quit
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

