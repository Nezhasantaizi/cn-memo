#!/bin/bash
cd "$(dirname "$0")"
source medanki_env/bin/activate
QT_QPA_PLATFORM=xcb python3 main.py.py