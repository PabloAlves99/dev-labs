from flask import Flask
from flask_restful import Api
from flask_jwt_extended import JWTManager

app = Flask(__name__)
api = Api(app)
jwt = JWTManager(app)


@app.route('/')
def hello_world():
    return 'Hello, World! This is a Flask API.'
