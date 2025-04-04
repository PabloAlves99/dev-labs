from flask import request
from flask_restful import Resource
from flask_jwt_extended import create_access_token
from models import User


class UserLogin(Resource):
    def post(self):
        data = request.json
        username = data.get("username")
        password = data.get("password")

        user = User.query.filter_by(username=username).first()
        if user and user.password == password:
            toekn = create_access_token(identity=username)
            return {"token": toekn}, 200

        return {"message": "Invalid credentials"}, 401
