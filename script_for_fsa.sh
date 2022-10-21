#!/bin/bash
#Обновление и загрузка пакетов
apt install sudo
sudo apt update
sudo apt upgrade
sudo apt -y install python3-pip
sudo apt -y install git
sudo apt -y install wget
sudo apt -y install libgtk-3-0
sudo apt -y install libasound2
sudo apt -y install libdbus-glib-1-2
sudo apt -y install libx11-xcb1
sudo apt -y install libgl1-mesa-glx
#
#Установка библиотек для ProgramFSA
pip3 install selenium
pip3 install openpyxl
#
#Создание необходимых папок
mkdir /opt/ProgramFSA
mkdir /opt/ProgramFSA/File
mkdir /opt/ProgramFSA/Screenshot
#
#Загрузка geckodriver
sudo wget https://github.com/mozilla/geckodriver/releases/download/v0.32.0/geckodriver-v0.32.0-linux64.tar.gz
tar xvzf geckodriver-v0.32.0-linux64.tar.gz
rm geckodriver-v0.32.0-linux64.tar.gz
mv geckodriver /opt/ProgramFSA
chmod 777 /opt/ProgramFSA/geckodriver
#
#Загрузка firefox
sudo wget https://ftp.mozilla.org/pub/firefox/releases/106.0/linux-x86_64/ru/firefox-106.0.tar.bz2
tar -x -j -f firefox-106.0.tar.bz2
rm firefox-106.0.tar.bz2
mv firefox /opt
ln -s /opt/firefox/firefox /usr/local/bin/firefox
chmod 777 -R /opt/firefox





