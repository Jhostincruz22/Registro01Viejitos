from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

# Nombre único del archivo centralizado de Excel
EXCEL_CENTRAL = str((BASE_DIR / "Años No registrados a partir de 1901.xlsx").resolve())

# Carpeta donde se almacenarán los respaldos automáticos
CARPETA_BACKUPS = str((BASE_DIR / "backups").resolve())
