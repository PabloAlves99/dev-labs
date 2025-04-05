from flask import request
from flask_restful import Resource
from database import db
from models import User


class UserRegister(Resource):
    def post(self):
        data = request.json
        username = data.get('username')
        password = data.get('password')
        email = data.get('email')

        if User.query.filter_by(username=username).first():
            return {'message': 'User already exists'}, 400

        new_user = User(username=username, password=password, email=email)
        db.session.add(new_user)
        db.session.commit()

        return {'message': 'User created successfully'}, 201
