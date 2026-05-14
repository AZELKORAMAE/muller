#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Extracteur de Fichiers Incorporés - Version Finale CORRIGÉE

CORRECTIONS APPLIQUÉES :
1. ✅ Activation Excel fonctionnelle (scripts VBS optimisés)
2. ✅ Fichiers ZIP remplacés dans documents modifiés
3. ✅ Nomenclature FJ corrigée (DER-348818_3_FJ_2 → FJ_3, FJ_4...)

Fonctionnalités:
- Détection et activation UNIQUEMENT des objets Excel embarqués
- Extraction avec ordre réel correct (PPTX, DOCX, XLSX)
- Remplacement direct (même nom que l'original)
- Indexation personnalisée via fichier de référence
- Rapport Excel détaillé
- Support: DOCX, DOC, XLSX, XLSM, XLS, PPTX, PPT, PDF, MSG
"""
import threading
import os
import sys
import shutil
import zipfile
import tempfile
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import subprocess
import xml.etree.ElementTree as ET
import re
import struct
import traceback
import time
try:
    from lxml import etree
except ImportError:
    install_package_safran("lxml")
    from lxml import etree
# Configuration repository Safran
SAFRAN_REPO = "--index-url https://artifacts.cloud.safran/repository/pypi-group/simple --trusted-host artifacts.cloud.safran"

def install_package_safran(package_name):
    """Installe un package depuis le repository Safran"""
    print(f"Installation de {package_name}...")
    cmd = f"pip install {package_name} {SAFRAN_REPO}"
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    return result.returncode == 0
