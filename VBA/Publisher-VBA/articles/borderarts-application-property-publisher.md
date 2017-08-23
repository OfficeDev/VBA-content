---
title: "Свойство BorderArts.Application (издатель)"
keywords: vbapb10.chm7733249
f1_keywords: vbapb10.chm7733249
ms.prod: publisher
api_name: Publisher.BorderArts.Application
ms.assetid: fc882feb-4e7d-f947-eaff-95c57cd2604e
ms.date: 06/08/2017
ms.openlocfilehash: 4b92738473a2f02a3ae6ea5d5d134142eaa5c863
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="borderartsapplication-property-publisher"></a>Свойство BorderArts.Application (издатель)

При использовании без квалификатор объекта, данное свойство возвращает объект **[приложения](application-object-publisher.md)** , который представляет текущего экземпляра Publisher. Используется квалификатор объекта, данное свойство возвращает объект **приложения** , представляющего создателя указанный объект. При использовании с помощью объекта OLE-автоматизации возвращает объект приложения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Приложения**

 переменная _expression_A, представляет собой объект- **BorderArts** .


## <a name="example"></a>Пример

В этом примере отображаются сведения о версии и построения для Publisher.


```vb
With Application 
 MsgBox "Current Publisher: version " _ 
 &; .Version &; " build " &; .Build 
End With
```

В этом примере отображается имя приложения, создавшего каждого связанного объекта на странице один активный публикации.




```vb
Dim shpOle As Shape 
 
For Each shpOle In ActiveDocument.Pages(1).Shapes 
 If shpOle.Type = pbLinkedOLEObject Then 
 MsgBox shpOle.OLEFormat.Application.Name 
 End If 
Next
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект BorderArts](borderarts-object-publisher.md)

