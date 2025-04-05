from datetime import datetime, timedelta, timezone
import jwt
from flask import request
from flask_restful import Resource
from models import User
from database import db
from config import Config


def generate_reset_token(user_id):
    payload = {
        "user_id": user_id,
        "exp": datetime.now(timezone.utc) + timedelta(minutes=15)
    }
    return jwt.encode(payload, Config.SECRET_KEY, algorithm="HS256")


def decode_reset_token(token):
    try:
        payload = jwt.decode(token, Config.SECRET_KEY, algorithms=["HS256"])
        return payload["user_id"]
    except jwt.ExpiredSignatureError:
        return None
    except jwt.InvalidTokenError:
        return None


class PasswordRecover(Resource):
    def post(self):
        data = request.get_json()
        email = data.get("email")
        user = User.query.filter_by(email=email).first()

        if user:
            token = generate_reset_token(user.id)
            link = f"http://localhost:5000/reset?token={token}"
            print("Link de recuperação:", link)

        return {"message": "Se o e-mail existir, o link foi enviado."}


class PasswordReset(Resource):
    def post(self):
        data = request.get_json()
        token = data.get("token")
        new_password = data.get("new_password")

        user_id = decode_reset_token(token)
        if not user_id:
            return {"message": "Token inválido ou expirado."}, 400

        user = User.query.get(user_id)
        if not user:
            return {"message": "Usuário não encontrado."}, 404

        user.password = new_password  # ou hash
        db.session.commit()

        return {"message": "Senha redefinida com sucesso!"}
