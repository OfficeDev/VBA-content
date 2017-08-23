---
title: "Свойство Application.ActiveWindow (издатель)"
keywords: vbapb10.chm131074
f1_keywords: vbapb10.chm131074
ms.prod: publisher
api_name: Publisher.Application.ActiveWindow
ms.assetid: 125e2bb4-f922-ceef-9e3e-5dbe3aaff2a4
ms.date: 06/08/2017
ms.openlocfilehash: 351070927e95ddf8b3e47d87ebb3fb67c34403bb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationactivewindow-property-publisher"></a>Свойство Application.ActiveWindow (издатель)

Возвращает объект **[Window](window-object-publisher.md)** , представляющий окно, фокус. Так как Microsoft Publisher имеет только одного окна, существует только один объект **Window** для возврата.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ActiveWindow**

 переменная _expression_A, представляющий объект **приложения** .


## <a name="example"></a>Пример

В этом примере отображается заголовка активного окна.


```vb
Sub CurrentCaption() 
 
 MsgBox ActiveDocument.ActiveWindow.Caption 
 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

