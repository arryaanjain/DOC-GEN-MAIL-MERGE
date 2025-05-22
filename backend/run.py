from flask import Flask
from flask_cors import CORS
from app.routes import routes

def create_app():
    app = Flask(__name__)
    app.config['UPLOAD_FOLDER'] = 'uploads/'
    # app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # Set as needed

    CORS(app)  # Enable CORS for all routes

    app.register_blueprint(routes, url_prefix='/api')

    return app

if __name__ == '__main__':
    app = create_app()
    app.run(debug=True)