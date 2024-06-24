import sys
import os

# Add your project directory to the sys.path
project_home = '/home/yourusername/yourproject'
if project_home not in sys.path:
    sys.path = [project_home] + sys.path

# Set the path to your virtualenv
activate_this = '/home/yourusername/.virtualenvs/myenv/bin/activate_this.py'
with open(activate_this) as file_:
    exec(file_.read(), dict(__file__=activate_this))

# Import your application
from myapp import app as application  # Adjust the import based on your app's structure
