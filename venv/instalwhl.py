# -*- coding: utf-8 -*-
import pip

def install_whl(path):
    pip.main(['install', path])

install_whl("D:\Tests\MyOffice_SDK_Document_API_Python_Win_2021.03_x64\MyOfficeSDKDocumentAPI-2021.3-cp38-cp38-win_amd64.whl")