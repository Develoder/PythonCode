# https://www.sites.google.com/site/mircurves/epicikloidy-i-gipocikloidy
import pygame
import math


class Main():
    # Глобальные переменные

    sizeDisplay = [300, 300]  # Размер окна
    colorDisplay = [255, 255, 255]  # Цвет окна

    center = [150, 150]  # Центер графики
    cellSize = [30, 30]  # Размер клетки

    FPS = 60  # ФПС

    # Проичий кал

    iteration = 0 # Итерация

    def __init__(self):
        pygame.init()

        self.scene = pygame.display.set_mode(self.sizeDisplay)
        self.scene.fill(self.colorDisplay)

        self.RuneGame()

    # Обнавление окна каждым кадром
    def RuneGame(self):
        clock = pygame.time.Clock()

        # Цикл игры
        running = True
        while running:
            self.scene.fill(self.colorDisplay)

            # Держим цикл на правильной скорости
            clock.tick(self.FPS)
            # Ввод процесса (события)
            for event in pygame.event.get():
                # проверяет закрытие окна
                if event.type == pygame.QUIT:
                    running = False

            # Обновление
            self.Update()

            # После отрисовки всего, переворачиваем экран
            pygame.display.flip()

        pygame.quit()

    # Функция вызывается каждый кадр
    def Update(self):
        self.GenerateGrid()

        self.iteration += 1  # Прогресс рисования (где 360 является концом круга)
        self.DrawEpizicloid(100, 10, 20, self.iteration)  # Генерация епициклоидной
        #self.DrawGipozicloid(100, 10, 20, self.iteration)  # Генерация гипоциклоидной

    # Вывод ЕПТциклоидной
    def DrawEpizicloid(self, R, s, d, iteration):
        if iteration > 360: # Проверка до завершении круга
            iteration = 360

        points = []
        for i in range(iteration + 1):
            t = 6.28 * i / 360
            x = self.center[0] + ((R + s) * math.cos(t) - d * math.cos((R + s) / s * t))
            y = self.center[1] + ((R + s) * math.sin(t) - d * math.sin((R + s) / s * t))
            points.append([x, y])

        pygame.draw.lines(self.scene, [185, 83, 94], False, points, 2)

    # Вывод гипоциклоиды
    def DrawGipozicloid(self, R, s, d, iteration):
        if iteration > 360:  # Проверка до завершении круга
            iteration = 360

        points = []
        for i in range(iteration + 1):
            t = 6.28 * i / 360
            x = self.center[0] + ((R - s) * math.cos(t) + d * math.cos((R - s) / s * t))
            y = self.center[1] + ((R - s) * math.sin(t) - d * math.sin((R - s) / s * t))
            points.append([x, y])

        pygame.draw.lines(self.scene, [185, 83, 94], False, points, 2)



    # Создает сетку
    def GenerateGrid(self):
        # Сетка
        gridX = []
        gridY = []
        ofsetGrid = [self.center[0] % self.cellSize[0], self.center[1] % self.cellSize[1]]
        for i in range(int(self.sizeDisplay[0] / self.cellSize[0]) + 2):
            gridX.append(i * self.cellSize[0] - ofsetGrid[0])

        for i in range(int(self.sizeDisplay[1] / self.cellSize[1]) + 2):
            gridY.append(i * self.cellSize[1] - ofsetGrid[1])

        for i in gridX:
            pygame.draw.line(self.scene, [101, 143, 163], [0, i], [self.sizeDisplay[0], i])

        for i in gridY:
            pygame.draw.line(self.scene, [101, 143, 163], [i, 0], [i, self.sizeDisplay[0]])




if __name__ == '__main__':
    main = Main()
