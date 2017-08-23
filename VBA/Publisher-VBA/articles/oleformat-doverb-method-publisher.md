---
title: "Метод OLEFormat.DoVerb (издатель)"
keywords: vbapb10.chm4456455
f1_keywords: vbapb10.chm4456455
ms.prod: publisher
api_name: Publisher.OLEFormat.DoVerb
ms.assetid: c4bca1f2-a3dd-0c49-1268-40e68e1fcef0
ms.date: 06/08/2017
ms.openlocfilehash: ff3431717629d21d84ba34907eecd90bd2b01906
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="oleformatdoverb-method-publisher"></a>Метод OLEFormat.DoVerb (издатель)

Запросы, что объект OLE выполните одно из его команд.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DoVerb** ( **_iVerb_**)

 переменная _expression_A, представляющий объект **OLEFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|iVerb|Обязательное свойство.| **Длинный**|Для выполнения команды. |

## <a name="remarks"></a>Заметки

Свойство **[ObjectVerbs](oleformat-objectverbs-property-publisher.md)** определяет доступные команды для объекта OLE.


## <a name="example"></a>Пример

В этом примере выполняется первая команда для третьего фигуры на первой странице active публикации, если фигуры — это связанный или внедренный объект OLE.


```vb
With ActiveDocument.Pages(1).Shapes(3) 
 If .Type = pbEmbeddedOLEObject Or _ 
 .Type = pbLinkedOLEObject Then 
 .OLEFormat.DoVerb (1) 
 End If 
End With
```

В этом примере выполняется команда «Открыть» для третьего фигуры на первой странице active публикации, если фигуры объекта OLE, который поддерживает команды «Открыть».




```vb
Dim strVerb As String 
Dim intVerb As Integer 
 
With ActiveDocument.Pages(1).Shapes(3) 
 
 ' Verify that the shape is an OLE object. 
 If .Type = pbEmbeddedOLEObject Or _ 
 .Type = pbLinkedOLEObject Then 
 
 ' Loop through the ObjectVerbs collection 
 ' until the "Open" verb is found. 
 For Each strVerb In .OLEFormat.ObjectVerbs 
 intVerb = intVerb + 1 
 If strVerb = "Open" Then 
 
 ' Perform the "Open" verb. 
 .OLEFormat.DoVerb iVerb:=intVerb 
 Exit For 
 End If 
 Next strVerb 
 End If 
End With 

```


