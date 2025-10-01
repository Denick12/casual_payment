class Config(object):
    DEBUG = False
    TESTING = False
    CSRF_ENABLED = True
    SECRET_KEY = 'This is the secret key'

    # Database Details
    DB_NAME = 'casual_payment'
    DB_USERNAME = 'root'
    DB_PASSWORD = ''
    DB_HOST = 'localhost'



class ProductionConfig(Config):
    pass


class DevelopmentConfig(Config):
    DEBUG = True
    SESSION_COOKIE_SECURE = True