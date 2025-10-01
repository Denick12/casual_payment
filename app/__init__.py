from flask import Flask

import config

app = Flask(__name__)

from app import users
app.config.from_object(config.DevelopmentConfig)