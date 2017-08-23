---
title: "Свойство PrintableRect.Application (издатель)"
keywords: vbapb10.chm7536642
f1_keywords: vbapb10.chm7536642
ms.prod: publisher
api_name: Publisher.PrintableRect.Application
ms.assetid: 8b035af1-75bd-bee6-8165-71c2f36bf839
ms.date: 06/08/2017
ms.openlocfilehash: 89cd9e4bb898ea31e51f8c0d47f8b56b54accd79
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="printablerectapplication-property-publisher"></a>Свойство PrintableRect.Application (издатель)

При использовании без квалификатор объекта, данное свойство возвращает объект **[приложения](application-object-publisher.md)** , который представляет текущего экземпляра Publisher. Используется квалификатор объекта, данное свойство возвращает объект **приложения** , представляющего создателя указанный объект. При использовании с помощью объекта OLE-автоматизации возвращает объект приложения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Приложения**

 переменная _expression_A, представляет собой объект- **PrintableRect** .


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


