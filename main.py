import asyncio
from CreateFiles import createFilesOzon
from Searchengines.startFind import OzonFind
async def main():
    price = 'result.xlsx'
    await OzonFind(price)  # Вставить файл с прайсом
    # print("Запуск 1 функции")
    await createFilesOzon(price) #Вставка пути к файлу Ozon
if __name__ == '__main__':
    asyncio.run(main())
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
