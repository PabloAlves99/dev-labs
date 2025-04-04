from flask import Flask
from flask_restful import Api
from flask_jwt_extended import JWTManager
from database import db
from resources.user import UserResource
from auth import UserLogin

app = Flask(__name__)
api = Api(app)
jwt = JWTManager(app)

db.init_app(app)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///example.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False


# Rotas
api.add_resource(UserResource, '/user/<int:user_id>')
api.add_resource(UserLogin, '/login')

if __name__ == '__main__':
    app.run(debug=True)
    with app.app_context():
        db.create_all()
