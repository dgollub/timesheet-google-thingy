#!/bin/bash

DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
ENV_FOLDER="${DIR}/env"

function abort() {
    echo $1
    echo "Aborted."
    exit 1;
}


which python > /dev/null 2>&1 || abort "python command not found"
which pip > /dev/null 2>&1 || abort "pip command not found"
which virtualenv > /dev/null 2>&1 || abort "virtualenv command not found"

if [ ! -d "${ENV_FOLDER}" ]; then
    virtualenv "${ENV_FOLDER}" || abort "Could not create folder with virtualenv: ${ENV_FOLDER}"
fi

. "${ENV_FOLDER}/bin/activate"

python "${DIR}/timesheet.py"

