#!/bin/bash
# 激活虚拟环境
source ./medank_env/bin/activate
# 设置QT平台，防止dxcb报错
export QT_QPA_PLATFORM=xcb
# 启动程序
python3 main.py
