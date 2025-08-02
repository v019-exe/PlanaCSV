#!/bin/bash

REPO_URL="https://github.com/v019-exe/PlanaCSV.git"
NOMBRE_APP="PlanaCSV_C"
ICONO_WIN="Plana.ico"
ICONO_UNIX="Plana.png"

necesita_root() {
    if [[ "$EUID" -ne 0 ]]; then
        echo "[!] Este paso requiere permisos de superusuario."
        sudo "$0" "$@"
        exit $?
    fi
}

existe_comando() {
    command -v "$1" >/dev/null 2>&1
}

verificar_o_instalar() {
    if ! existe_comando "$1"; then
        echo "[!] '$1' no encontrado. Instalando..."
        case "$SISTEMA" in
            Linux)
                if existe_comando apt; then
                    necesita_root
                    sudo apt update && sudo apt install -y "$2"
                elif existe_comando dnf; then
                    necesita_root
                    sudo dnf install -y "$2"
                fi
                ;;
            Darwin)
                if ! existe_comando brew; then
                    echo "Instala Homebrew primero: https://brew.sh"
                    exit 1
                fi
                brew install "$2"
                ;;
            MINGW*|MSYS*|CYGWIN*)
                echo "[!] Instala '$1' manualmente en Windows."
                ;;
        esac
    fi
}

SISTEMA="$(uname)"
echo "[+] Sistema detectado: $SISTEMA"

verificar_o_instalar python3 python3
verificar_o_instalar pip3 python3-pip
verificar_o_instalar git git
verificar_o_instalar gcc gcc
verificar_o_instalar make make

echo "[+] Clonando el repositorio..."
git clone "$REPO_URL" "$NOMBRE_APP-src" || exit 1
cd "$NOMBRE_APP-src" || exit 1

python3 -m venv venv
source venv/bin/activate 2>/dev/null || source venv/Scripts/activate

echo "[+] Instalando Nuitka..."
pip install --upgrade pip
pip install nuitka

if [[ "$SISTEMA" == "Linux" ]]; then
    nuitka main.py \
        --onefile \
        --standalone \
        --enable-plugin=tk-inter \
        --include-data-files=$ICONO_UNIX=$ICONO_UNIX \
        --output-filename=$NOMBRE_APP

    mkdir -p ~/.local/bin/$NOMBRE_APP
    cp $NOMBRE_APP ~/.local/bin/$NOMBRE_APP/
    cp $ICONO_UNIX ~/.local/bin/$NOMBRE_APP/

    cat <<EOF > ~/.local/share/applications/${NOMBRE_APP}.desktop
[Desktop Entry]
Type=Application
Name=PlanaCSV
Exec=$HOME/.local/bin/$NOMBRE_APP/$NOMBRE_APP
Icon=$HOME/.local/bin/$NOMBRE_APP/$ICONO_UNIX
Terminal=false
Categories=Utility;
EOF

    chmod +x ~/.local/share/applications/${NOMBRE_APP}.desktop
    echo "[✓] Compilación y acceso directo listos en Linux."

elif [[ "$SISTEMA" == "Darwin" ]]; then
    nuitka main.py \
        --onefile \
        --standalone \
        --enable-plugin=tk-inter \
        --include-data-files=$ICONO_UNIX=$ICONO_UNIX \
        --output-filename=$NOMBRE_APP

    echo "[✓] Compilación para macOS lista."

elif [[ "$SISTEMA" == MINGW* || "$SISTEMA" == MSYS* || "$SISTEMA" == CYGWIN* ]]; then
    nuitka main.py \
        --onefile \
        --standalone \
        --windows-console-mode=disable \
        --assume-yes-for-downloads \
        --enable-plugin=tk-inter \
        --windows-icon-from-ico=$ICONO_WIN \
        --include-data-files=$ICONO_WIN=$ICONO_WIN \
        --output-filename=$NOMBRE_APP.exe

    echo "[✓] Compilación para Windows lista: $NOMBRE_APP.exe"
else
    echo "[!] Sistema operativo no compatible automáticamente."
    exit 1
fi