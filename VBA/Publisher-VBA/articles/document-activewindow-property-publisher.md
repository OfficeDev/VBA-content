---
title: "Свойство Document.ActiveWindow (издатель)"
keywords: vbapb10.chm196611
f1_keywords: vbapb10.chm196611
ms.prod: publisher
api_name: Publisher.Document.ActiveWindow
ms.assetid: 0d00a8fa-aef2-43df-3c54-0cca804b7eee
ms.date: 06/08/2017
ms.openlocfilehash: 6d4e33cafa90a9c2a9a9a234bbca05567e84678e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentactivewindow-property-publisher"></a>Свойство Document.ActiveWindow (издатель)

Возвращает объект **[Window](window-object-publisher.md)** , представляющий окно, фокус. Так как Microsoft Publisher имеет только одного окна, существует только один объект **Window** для возврата.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ActiveWindow**

 переменная _expression_A, представляющий объект **Document** .


## <a name="example"></a>Пример

В этом примере отображается заголовка активного окна.


```vb
Sub CurrentCaption() 
 
 MsgBox ActiveDocument.ActiveWindow.Caption 
 
End Sub
```


