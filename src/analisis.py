from enum import Enum, unique

@unique
class Analisis(Enum):
    dca2f = 'dca2f'
    dbca1f = 'dbca1f'
    dca1f = 'dca1f'

    @classmethod
    def has_value(cls, value):
        return value in [item.value for item in cls]

    @classmethod
    def get_item(cls, value):
        for item in cls:
            if value == item.value:
                return item


if __name__ == '__main__':
    print(Analisis.dca2f in Analisis)
    print(Analisis.has_value('dca2f'))
    print(Analisis.get_item('dca2f'))
    print(type(Analisis.get_item('dca2f')))