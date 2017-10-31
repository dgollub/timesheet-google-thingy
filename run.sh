#!/bin/bash

DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
ENV_FOLDER="${DIR}/env"
REQUIREMENTS_FILE="${DIR}/requirements.txt"

function abort() {
    echo $1
    echo "Aborted."
    exit 1;
}


which python > /dev/null 2>&1 || abort "python command not found"
which pip > /dev/null 2>&1 || abort "pip command not found"
which virtualenv > /dev/null 2>&1 || abort "virtualenv command not found"

if [ ! -d "${ENV_FOLDER}" ]; then
    if [ ! -f "${REQUIREMENTS_FILE}" ]; then
        abort "The ${REQUIREMENTS_FILE} file for pip could not be found."
    fi
    virtualenv "${ENV_FOLDER}" || abort "Could not create folder with virtualenv: ${ENV_FOLDER}"
    . "${ENV_FOLDER}/bin/activate"
    pip install -r "${REQUIREMENTS_FILE}" || abort "Could not pip install all requirements. Please check the output above."
else
    . "${ENV_FOLDER}/bin/activate"
fi

python "${DIR}/timesheet.py" "$1"
