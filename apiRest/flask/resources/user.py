from flask_restful import Resource
from models import User


class UserResource(Resource):

    def get(self, user_id):
        user = User.query.get(user_id)
        if user:
            return {"id": user.id, "username": user.username}, 200
        return {"message": "User not found"}, 404
