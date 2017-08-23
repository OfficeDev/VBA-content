---
title: "Свойство Hyperlink.TargetType (издатель)"
keywords: vbapb10.chm4587529
f1_keywords: vbapb10.chm4587529
ms.prod: publisher
api_name: Publisher.Hyperlink.TargetType
ms.assetid: 1cbc8c36-563c-4464-4f0d-2836682ce532
ms.date: 06/08/2017
ms.openlocfilehash: f642d968b0df60549c6ce36a5b7c81f1ea5e0a6e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="hyperlinktargettype-property-publisher"></a>Свойство Hyperlink.TargetType (издатель)

Возвращает константу **PbHlinkTargetType** , представляющий тип гиперссылки. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TargetType**

 переменная _expression_A, представляющий объект **гиперссылки** .


### <a name="return-value"></a>Возвращаемое значение

PbHlinkTargetType


## <a name="remarks"></a>Заметки

Значение свойства **TargetType** может иметь одно из следующих констант **PbHlinkTargetType** .



| **pbHlinkTargetTypeEmail**|| **pbHlinkTargetTypeFirstPage**|| **pbHlinkTargetTypeLastPage**|| **pbHlinkTargetTypeNextPage**|| **pbHlinkTargetTypeNone**|| **pbHlinkTargetTypePageID**|| **pbHlinkTargetTypePreviousPage**|| **pbHlinkTargetTypeURL**|

## <a name="example"></a>Пример

В этом примере выполняется проверка, что указанный гиперссылка — это URL-адрес и если он установлен, задает текст гиперссылки и адрес. В этом примере предполагается, что имеется по крайней мере один фигуры на первой странице active публикации.


```vb
Sub SetHyperlinkTextToDisplay() 
 With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Hyperlinks.Item(1) 
 If .TargetType = pbHlinkTargetTypeURL Then 
 .TextToDisplay = "Tailspin Toys Web Site" 
 .Address = "http://www.tailspintoys.com/" 
 End If 
 End With 
End Sub
```


