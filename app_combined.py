from app import app as v28_app
from app_staff import app as staff_app
from werkzeug.middleware.dispatcher import DispatcherMiddleware
from flask import Flask

app = Flask(__name__)

app.wsgi_app = DispatcherMiddleware(v28_app, {
    '/staff': staff_app
})
