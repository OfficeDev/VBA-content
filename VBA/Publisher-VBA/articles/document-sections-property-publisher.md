---
title: "Свойство Document.Sections (издатель)"
keywords: vbapb10.chm196738
f1_keywords: vbapb10.chm196738
ms.prod: publisher
api_name: Publisher.Document.Sections
ms.assetid: 9e425836-1d62-99ef-2984-b61f3a3cf831
ms.date: 06/08/2017
ms.openlocfilehash: 5546ce18edfe290729d4defb69ab4b524278ca67
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentsections-property-publisher"></a>Свойство Document.Sections (издатель)

Возвращает объект, который **разделах** представляет коллекцию объектов **раздела** в указанный документ. Только для чтения, **разделы**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Разделы**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

Разделы


## <a name="example"></a>Пример

В этом примере задается объектную переменную объекту **разделах** активного документа и добавляет новый раздел, начиная с вторая страница публикации. В этом примере предполагается, что в публикации есть по крайней мере две страницы.


```vb
Dim objSections As Sections 
Set objSections = ActiveDocument.Sections 
objSections.Add StartPageIndex:=2 

```


