import pygame
from random import randint
import os


class BallSettings:
    def __init__(self) -> None:
        self.ball_settings()

    def ball_settings(self):
        self.ball_size = 25
        self.ball = pygame.Rect(randint(5, 500), 500,
                                self.ball_size, self.ball_size)
        self.ball_movement = [3, -3]

        current_dir = os.path.dirname(__file__)
        self.ball_image = pygame.image.load(
            os.path.join(current_dir, 'images', 'ball.png'))
        self.ball_image = pygame.transform.scale(
            self.ball_image, (self.ball_size, self.ball_size))
