import sys
from cx_Freeze import setup, Executable

# Dépendances
build_exe_options = {
    "packages": ["os", "pandas", "openpyxl", "PIL"],
    "include_files": ["logo_diginamic.png"],
}

# Créer l'exécutable
setup(
    name="NomDeVotreApplication",
    version="0.1",
    description="Description de votre application",
    options={"build_exe": build_exe_options},
    executables=[Executable("PythonXlsx.py", base=None)]
)