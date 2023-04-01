from setuptools import setup
from pathlib import Path

with Path(__file__).parent.joinpath("README.md").open(encoding="utf-8") as f:
    long_description = f.read()

setup(
    name="conversor_nominas_bancos_chile",
    version="1.7.7",
    description="Librería que convierte el formato de nóminas del BCI al formato del resto de bancos.",
    author="Antonio Canada Momblant",
    author_email="toni.cm@gmail.com",
    packages=['conversor_nominas_bancos_chile'],
    package_data={'conversor_nominas_bancos_chile': [
        'bancos_codigos.json', 'bancos_headers_nomina.json']},
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    install_requires=[
        'pandas',
        'numpy',
        'datetime',
        'pathlib',
        'tk',
        'openpyxl',
        'xlrd'],
    entry_points={
        'console_scripts': [
            'start_menu_conversor_nominas = conversor_nominas_bancos_chile.bank_tkinter_menu:iniciar_menu'
        ]
    },
    long_description=long_description,
)
