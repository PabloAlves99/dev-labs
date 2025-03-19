import pygame
import os
from player_settings import PlayerSettings
from ball_settings import BallSettings


class GameSettings():
    def __init__(self) -> None:
        self.initial_screen()
        self.initial_blocks()
        self.color_settings()
        self.player_settings = PlayerSettings()
        self.ball_settings = BallSettings()
        self.blocks = self.blocks_settings()

    def initial_screen(self):
        self.screen_size = (620, 600)
        self.screen = pygame.display.set_mode(self.screen_size)
        pygame.display.set_caption("Brick Breaker")

    def initial_blocks(self):
        self.num_of_blocks_in_line = 8
        self.num_of_block_rows = 10
        self.total_number_of_blocks = self.num_of_blocks_in_line * \
            self.num_of_block_rows

    def color_settings(self):
        self.colors = {
            "background": "#BF9D7E",
            "punctuation": "#26130D",
            "player_color": "#26130D",
            "special_block": "#FFD700",
        }

    def blocks_settings(self):
        self.distance_between_blocks = 5
        self.block_width = self.screen_size[0] / 8 - \
            self.distance_between_blocks
        self.block_height = 25
        self.distance_between_rows = self.block_height + 5
        self._blocks = []

        current_dir = os.path.dirname(__file__)
        self.block_image = pygame.image.load(
            os.path.join(current_dir, 'images', 'block-100.png'))
        self.block_image = pygame.transform.scale(
            self.block_image, (self.block_width, self.block_height))

        return self.generate_blocks()

    def generate_blocks(self):

        for j in range(self.num_of_block_rows):

            for i in range(self.num_of_blocks_in_line):

                block = pygame.Rect(
                    3 + i * (self.block_width + self.distance_between_blocks),
                    3 + j * self.distance_between_rows, self.block_width,
                    self.block_height)
                self._blocks.append(block)

        return self._blocks

    def update_score(self, punctuation):
        font = pygame.font.Font(None, 30)
        text = font.render(f"Pontuação: {punctuation}", 1,
                           self.colors["punctuation"])
        self.screen.fill(
            self.colors["background"], (0, 580, 150, 30))
        self.screen.blit(text, (0, 580))
        if punctuation >= self.total_number_of_blocks:
            return True
        else:
            return False
