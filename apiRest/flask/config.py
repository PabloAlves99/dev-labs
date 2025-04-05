class Config:
    SECRET_KEY = "minha_chave_super_secreta"
    JWT_SECRET_KEY = "minha_chave_jwt"
    SQLALCHEMY_DATABASE_URI = "sqlite:///database.db"
    SQLALCHEMY_TRACK_MODIFICATIONS = False
