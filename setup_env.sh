#!/bin/bash

VENV_NAME="venv"
PATH=$(pwd)

# define interpreter
read -p "Declare python interpreter for packages [default: /Users/nikhildatta/anaconda3/bin/python]: " PYTHON_ALIAS
if [[ -z "$PYTHON_ALIAS"]]; then
    PYTHON_ALIAS="/Users/nikhildatta/anaconda3/bin/python"
fi

PYTHON_FOLDER=$(echo "$PYTHON_ALIAS" | awk -F'/' '{print $(NF-1) "/" $NF}')

# define venv details
read -p "Enter virtual environment name [default: venv]" VENV_NAME
if [[ -z "$VENV_NAME" ]]; then
    VENV_NAME="venv"
fi

echo "Python Interpreter: $PYTHON_ALIAS"
echo "Python Folder Struct: $PYTHON_FOLDER"
echo "Virtual Environment: $VENV_NAME"

# create venv
$PYTHON_ALIAS -m venv --system-site-packages $VENV_NAME

# activate venv
source $VENV_NAME/bin/activate

# check venv is active
if [[ $(which python) != $(pwd)/$VENV_NAME/$PYTHON_FOLDER ]]; then
    echo "Failed to activate venv"
    exit 1
else
    echo "venv created - setting as alias"
    alias python=$(pwd)/$VENV_NAME/$PYTHON_FOLDER
fi

# create test.env
cat <<EOL > $(pwd)/environments/.env
ENV_PATH_PREFIX=$(pwd)
PYTHON_PATH=$(pwd)/$VENV_NAME/$PYTHON_FOLDER
EOL
echo ".env created"

# create default.py
DEF_FILE="$(pwd)/utils/default.py"

mkdir -p "$(dirname "$DEF_FILE")"

cat<<EOL >"$DEF_FILE"
from dotenv import load_dotenv
import configparser

def load_env(env_path="$(pwd)/environments/.env"):
    "Load env vars from .env file"
    load_dotenv(env_path)

def load_cfg(config_path="$(pwd)/utils/config.ini"):
    "Load configurations from .ini file"
    config = configparser.ConfigParser()
    config.read(config_path)
    return config
EOL

echo "default.py created"

# install local packages
pip install -e .
if [[ $? -ne 0 ]]; then
    echo "packages not installed"
    false
fi
echo "local packages installed to venv"

export PYTHONPATH=$(pwd)

echo "setup complete"