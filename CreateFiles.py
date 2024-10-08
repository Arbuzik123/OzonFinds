import math
import chromedriver_autoinstaller
from multiprocessing import Process, Lock, Semaphore
import pandas as pd
import os
from DefOzon.updateOzon import updateOzon
chromedriver_autoinstaller.install()
drivers = []
async def createFilesOzon(cat):
    df = pd.read_excel(cat)
    total_rows = len(df)
    num_chunks = min(total_rows, 1)
    chunk_size = math.ceil(total_rows / num_chunks)
    chunks = [df[i:i + chunk_size] for i in range(0, total_rows, chunk_size)]
    if len(chunks) < num_chunks:
        chunks.extend([pd.DataFrame()] * (num_chunks - len(chunks)))
    for idx, chunk in enumerate(chunks):
        path = rf"UpdatesOzon/" + cat.replace(".xlsx", "") + f"_{idx + 1}.xlsx"
        chunk.to_excel(rf'{path}', index=False)
    # Создаем список задач
    processes = []
    lock = Semaphore(0)
    window_PosY = 1080 / 2
    window_PosX = 1920 / 2
    screen_width = 1920
    screen_height = 1090
    window_width = screen_width // 2  # Ширина окна - половина экрана
    window_height = screen_height // 2  # Высота окна - половина экрана
    positions = [
        (screen_width // 2, screen_height // 2),
        # Позиция первого окна (левый верхний угол)
        (screen_width // 2, 0),  # Позиция второго окна (правый верхний угол)
        (0, screen_height // 2),
        (0, 0)  # Позиция третьего окна (левый нижний угол)
        # Позиция четвертого окна (правый нижний угол)
    ]
    for i in range(num_chunks):
        p = Process(target=updateOzon, args=(i, path, lock, window_PosX, window_PosY, positions.pop(0)))
        processes.append(p)
    # Запускаем первый процесс и блокируем остальные
    processes[0].start()
    for p in processes[1:]:
        lock.acquire()
        p.start()
    for p in processes:
        p.join()
    input_folder = 'UpdatesOzon'
    # Получаем список файлов Excel в папке
    files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx')]
    # Создаем пустой DataFrame, в который будем добавлять данные
    result_df = pd.DataFrame()
    # Читаем каждый файл и добавляем его содержимое к результату
    for file in files:
        df = pd.read_excel(os.path.join(input_folder, file))
        result_df = pd.concat([result_df, df])
    # Путь для сохранения результирующего файла Excel
    output_path = os.path.join('UpdateResultOzon', rf'{cat.replace(".xlsx", "")}.xlsx')
    # Сохраняем результирующий DataFrame в файл
    result_df.to_excel(output_path, index=False)