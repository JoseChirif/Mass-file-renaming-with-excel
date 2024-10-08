from setuptools import setup, find_packages
from src import __version__ as version  # Import version from src/__init__.py

#Requirements
with open('requirements.txt') as f:
    install_requires = list(map(lambda x: x.strip(), f.readlines()))
    
    
setup(
    name='Mass file renaming',
    version=version, 
    description='Automatically creates an excel file with all files in the same folder and rename those files with the new value that you assign.',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    author='Jose Chirif',
    author_email='',
    url='https://github.com/tu_usuario/nombre_del_paquete',
    packages=find_packages(where='src'),
    package_dir={'': 'src'},  # Especifica que los paquetes están en 'src'
    install_requires=install_requires,  # Usa la lista de dependencias
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Intended Audience :: End Users/Desktop',
        "Operating System :: Microsoft :: Windows",
    ],
    python_requires='>=3.11.5',
    include_package_data=True,
    data_files=[
        ('', ['run.py', 'LICENSE.txt', 'README.md']),  # Archivos en el directorio raíz
        ('assets', ['assets/icon.ico', 'assets/icon.png', 'icon_github.png']),  # Archivos en la carpeta assets
        ('config', ['config/config.py']),  # Archivo en la carpeta config
        ('functions', ['functions/functions.py']),  # Archivos en functions
        ('locales', ['locales/en.json', 'locales/es.json', 'languages.json']),  # Archivos en locales
        ('src', ['src/a_Importar_archivos_a_excel.py', 'src/b_Modificar_nombres.py', 'c_Desbloquear_excel.py', 'main_menu.py']),  # Archivos en src
    ],
     entry_points={
        'console_scripts': [
            'Mass_files_renaming = run:main',  # 'mi_comando' es el comando que se ejecutará en la línea de comandos
        ],
    },
)