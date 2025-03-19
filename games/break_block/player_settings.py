import pygame


class PlayerSettings:
    def __init__(self) -> None:
        self.player_settings()

    def player_settings(self):
        self.player_size = 150
        self.player = pygame.Rect(300, 550, self.player_size, 15)
