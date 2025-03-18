# pylint: disable=all
import os
import random
import pygame


class GameSettings:
    def __init__(self) -> None:
        self.screen_settings()
        self.upload_images()

    def screen_settings(self):
        self.screen_width = 500
        self.screen_height = 800
        pygame.font.init()
        self.dots_font = pygame.font.SysFont('arial', 50)

    def upload_images(self):
        current_dir = os.path.dirname(__file__)
        self._img_pipe = pygame.transform.scale2x(
            pygame.image.load(os.path.join(current_dir, 'imgs', 'pipe.png')))
        self._img_floor = pygame.transform.scale2x(
            pygame.image.load(os.path.join(current_dir, 'imgs', 'base.png')))
        self._img_background = pygame.transform.scale2x(
            pygame.image.load(os.path.join(current_dir, 'imgs', 'bg.png')))
        self._img_bird = [
            pygame.transform.scale2x(pygame.image.load(
                os.path.join(current_dir, 'imgs', 'bird1.png'))),
            pygame.transform.scale2x(pygame.image.load(
                os.path.join(current_dir, 'imgs', 'bird2.png'))),
            pygame.transform.scale2x(pygame.image.load(
                os.path.join(current_dir, 'imgs', 'bird3.png'))),
        ]


class Bird(GameSettings):

    def __init__(self, x, y):
        super().__init__()
        self.bird_position(x, y)
        self.bird_settings()

    def bird_settings(self):
        self.bird_maximum_rotation = 25
        self.bird_rotation_speed = 20
        self.bird_animation_time = 5
        self.bird_angle = 0
        self.bird_speed = 0
        self.bird_height = self.bird_y
        self.bird_time = 0
        self.bird_image_count = 0
        self.bird_image = self._img_bird[0]

    def bird_position(self, x, y):
        self.bird_x = x
        self.bird_y = y

    def bird_jump(self):
        self.bird_speed = -10.5
        self.bird_time = 0
        self.bird_height = self.bird_y

    def bird_move(self):
        self.bird_time += 1
        self.bird_displacement = 1.5 * (self.bird_time**2) + \
            self.bird_speed * self.bird_time

        if self.bird_displacement > 16:
            self.bird_displacement = 16

        elif self.bird_displacement < 0:
            self.bird_displacement -= 2

        self.bird_y += self.bird_displacement

        if self.bird_displacement < 0 or self.bird_y < (self.bird_height + 50):
            if self.bird_angle < self.bird_maximum_rotation:
                self.bird_angle = self.bird_maximum_rotation
        else:
            if self.bird_angle > -90:
                self.bird_angle -= self.bird_rotation_speed

    def bird_design(self, screen):
        self.bird_image_count += 1

        if self.bird_image_count < self.bird_animation_time:
            self.bird_imagem = self._img_bird[0]
        elif self.bird_image_count < self.bird_animation_time * 2:
            self.bird_imagem = self._img_bird[1]
        elif self.bird_image_count < self.bird_animation_time * 3:
            self.bird_imagem = self._img_bird[2]
        elif self.bird_image_count < self.bird_animation_time * 4:
            self.bird_imagem = self._img_bird[1]
        elif self.bird_image_count >= self.bird_animation_time * 4 + 1:
            self.bird_imagem = self._img_bird[0]
            self.bird_image_count = 0

        if self.bird_angle <= -80:
            self.bird_image = self._img_bird[1]
            self.bird_image_count = self.bird_animation_time * 2

        rotated_image = pygame.transform.rotate(
            self.bird_image, self.bird_angle)
        image_center_position = self.bird_imagem.get_rect(
            topleft=(self.bird_x, self.bird_y)).center
        rectangle = rotated_image.get_rect(center=image_center_position)
        screen.blit(rotated_image, rectangle.topleft)

    def get_mask(self):
        return pygame.mask.from_surface(self.bird_image)


class Pipe(GameSettings):

    def __init__(self, x):
        super().__init__()
        self.pipe_x = x
        self.pipe_height = 0
        self.pipe_top_position = 0
        self.pipe_base_position = 0
        self.pipe_top = pygame.transform.flip(self._img_pipe, False, True)
        self.pipe_base = self._img_pipe
        self.pipe_surpassed = False
        self.pipe_distance = 200
        self.pipe_speed_screen = 5
        self.pipe_set_height()

    def pipe_set_height(self):
        self.pipe_height = random.randrange(50, 450)
        self.pipe_top_position = self.pipe_height - self.pipe_top.get_height()
        self.pipe_base_position = self.pipe_height + self.pipe_distance

    def pipe_move(self):
        self.pipe_x -= self.pipe_speed_screen

    def pipe_design(self, tela):
        tela.blit(self.pipe_top, (self.pipe_x, self.pipe_top_position))
        tela.blit(self.pipe_base, (self.pipe_x, self.pipe_base_position))

    def colide(self, bird):
        bird_mask = bird.get_mask()
        top_mask = pygame.mask.from_surface(self.pipe_top)
        base_mask = pygame.mask.from_surface(self.pipe_base)

        distance_top = (self.pipe_x - bird.bird_x,
                        self.pipe_top_position - round(bird.bird_y))
        distance_base = (self.pipe_x - bird.bird_x,
                         self.pipe_base_position - round(bird.bird_y))

        top_point = bird_mask.overlap(top_mask, distance_top)
        base_point = bird_mask.overlap(base_mask, distance_base)

        if base_point or top_point:
            return True
        else:
            return False


class Floor(GameSettings):

    def __init__(self, y):
        super().__init__()
        self.floor_width = self._img_floor.get_width()
        self.flor_image = self._img_floor
        self.floor_y = y
        self.floor_x1 = 0
        self.floor_x2 = self.floor_width
        self.floor_speed = 5

    def floor_move(self):
        self.floor_x1 -= self.floor_speed
        self.floor_x2 -= self.floor_speed

        if self.floor_x1 + self.floor_width < 0:
            self.floor_x1 = self.floor_x2 + self.floor_width
        if self.floor_x2 + self.floor_width < 0:
            self.floor_x2 = self.floor_x1 + self.floor_width

    def floor_design(self, screen):
        screen.blit(self.flor_image, (self.floor_x1, self.floor_y))
        screen.blit(self.flor_image, (self.floor_x2, self.floor_y))


class GameManager(GameSettings):
    def __init__(self) -> None:
        super().__init__()

    def draw_screen(self, screen, birds, pipes, floor, points):
        screen.blit(self._img_background, (0, 0))
        for bird in birds:
            bird.bird_design(screen)
        for pipe in pipes:
            pipe.pipe_design(screen)

        self.text = self.dots_font.render(
            f"Pontuação: {points}", 1, (255, 255, 255))
        screen.blit(self.text, (self.screen_width -
                    10 - self.text.get_width(), 10))
        floor.floor_design(screen)
        pygame.display.update()


def main():
    birds = [Bird(230, 350)]
    floor = Floor(730)
    pipe = [Pipe(700)]
    screen = pygame.display.set_mode((500, 800))
    points = 0
    clock = pygame.time.Clock()

    running = True
    while running:
        clock.tick(30)

        for evento in pygame.event.get():
            if evento.type == pygame.QUIT:
                running = False
                pygame.quit()
                quit()
            if evento.type == pygame.KEYDOWN:
                if evento.key == pygame.K_SPACE:
                    for bird in birds:
                        bird.bird_jump()

        for bird in birds:
            bird.bird_move()
        floor.floor_move()

        add_pipe = False
        remove_pipe = []
        for _pipe in pipe:
            for i, bird in enumerate(birds):
                if _pipe.colide(bird):
                    birds.pop(i)
                    pygame.quit()
                    quit()
                if not _pipe.pipe_surpassed and bird.bird_x > _pipe.pipe_x:
                    _pipe.pipe_surpassed = True
                    add_pipe = True
            _pipe.pipe_move()
            if _pipe.pipe_x + _pipe.pipe_top.get_width() < 0:
                remove_pipe.append(_pipe)

        if add_pipe:
            points += 1
            pipe.append(Pipe(600))
        for _pipe in remove_pipe:
            pipe.remove(_pipe)

        for i, bird in enumerate(birds):
            if (bird.bird_y + bird.bird_image.get_height()) >\
                    floor.floor_y or bird.bird_y < 0:
                birds.pop(i)
                pygame.quit()
                quit()

        game = GameManager()
        game.draw_screen(screen, birds, pipe, floor, points)


if __name__ == '__main__':
    main()
