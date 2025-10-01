from app import app
import pymysql
from flask_wtf.csrf import CSRFProtect
import pandas as pd

csrf = CSRFProtect(app)


def db_connection():
    conn = pymysql.connect(host=app.config["DB_HOST"], user=app.config["DB_USERNAME"],
                           password=app.config["DB_PASSWORD"],
                           database=app.config["DB_NAME"])
    cursor = conn.cursor()
    return conn, cursor


def allowed_file(filename):
    return ('.' in filename and
            filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS'])

