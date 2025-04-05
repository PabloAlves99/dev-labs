from flask import Flask
from flask_restful import Api
from flask_jwt_extended import JWTManager
from database import db
from resources.user import UserResource
from resources.register import UserRegister
from auth import UserLogin
from config import Config

app = Flask(__name__)
app.config.from_object(Config)

db.init_app(app)
api = Api(app)
jwt = JWTManager(app)

# Rotas
api.add_resource(UserRegister, '/register')
api.add_resource(UserLogin, '/login')
api.add_resource(UserResource, '/user/<int:user_id>')

if __name__ == '__main__':
    app.run(debug=True)
    with app.app_context():
        db.create_all()
