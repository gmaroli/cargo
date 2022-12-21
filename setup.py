from cx_Freeze import setup, Executable

base = None

executables = [Executable("extract_manifest_data.py", base=base)]

packages = ["idna", 'pandas', 'openpyxl', 'datetime', 'os']
options = {
    'build_exe': {
        'packages': packages,
    },
}

setup(
    name="Create_Manifest",
    options=options,
    version="1.0",
    description='test description',
    executables=executables
)
