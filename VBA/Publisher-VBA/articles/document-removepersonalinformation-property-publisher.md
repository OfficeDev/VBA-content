---
title: "Свойство Document.RemovePersonalInformation (издатель)"
keywords: vbapb10.chm196742
f1_keywords: vbapb10.chm196742
ms.prod: publisher
api_name: Publisher.Document.RemovePersonalInformation
ms.assetid: bbc1aee1-90ca-966e-c17c-579064318cd1
ms.date: 06/08/2017
ms.openlocfilehash: 927cdf34920c0dfc208c1e9ec4032485b5d476d1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentremovepersonalinformation-property-publisher"></a>Свойство Document.RemovePersonalInformation (издатель)

Возвращает или задает **логическое** , представляет ли сохраняться личные сведения при сохранении файла. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RemovePersonalInformation**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Информация, удалены из документа является автор, руководитель, компании и идентификатор GUID для компьютера, на котором был создан документ.

По умолчанию для этого свойства имеет **значение False**.


## <a name="example"></a>Пример

В этом примере удаляется личных сведений из активных документов.


```vb
ActiveDocument.RemovePersonalInformation = True 

```


