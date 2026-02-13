import random

def Main():
    test = [    'Режим ожидания',
                'Готово',
                'Пожалуйста, подождите',
                'Обработка',
                'Подождите', ]
    iValue = random.randint(0, len(test)-1)
    # iValue = 1
    szValue = test[iValue]

    print(f'${Reverse(szValue)} {szValue}')

def Reverse(szString):
    match szString:
        case 'Режим ожидания' | 'Готово':
            return 'Норма'
        case 'Пожалуйста, подождите' | 'Обработка' | 'Подождите':
            return 'В работе'  
    return 'Неизвестно'

Main()