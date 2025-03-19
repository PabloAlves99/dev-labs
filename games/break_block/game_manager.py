import pygame
from actions import Actions


class GameManager(Actions):
    def __init__(self) -> None:
        super().__init__()
        self.end_game = False
        self.punctuation = 0
        self.event = ...

    def draw_home_screen(self):
        self.screen.fill(self.colors["background"])
        pygame.draw.rect(self.screen,
                         self.colors["player_color"],
                         self.player_settings.player)

        self.screen.blit(self.ball_settings.ball_image,
                         self.ball_settings.ball)

    def draw_blocks(self):
        for block in self.blocks:
            self.screen.blit(self.block_image, block.topleft)

    def play(self):
        self.draw_home_screen()
        self.draw_blocks()
        self.end_game = self.update_score(
            self.total_number_of_blocks - len(self.blocks))
        pygame.display.flip()
        try:
            for _event in pygame.event.get():
                self.event = _event
                if self.event.type == pygame.QUIT:
                    self.end_game = True
            self.player_move(self.event)
            __ball_move = self.ball_move()

            if not __ball_move:
                self.end_game = True

            pygame.time.wait(10)
        except AttributeError:
            pass
