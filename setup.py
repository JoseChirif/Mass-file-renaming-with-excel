from setuptools import setup, find_packages
from config.config import __version__

#Requirements
with open('requirements.txt') as f:
    install_requires = list(map(lambda x: x.strip(), f.readlines()))
    
    
setup(
    name='Mass file renaming',
    version=__version__, 
    description='Automatically creates an excel file with all files in the same folder and rename those files with the new value that you assign.',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    author='Jose Chirif',
    url='https://github.com/JoseChirif/Mass-file-renaming-with-excel',
    packages=find_packages(where='src'),
    package_dir={'': 'src'},  
    install_requires=install_requires,  
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Intended Audience :: End Users/Desktop',
        "Operating System :: Microsoft :: Windows",
    ],
    python_requires='>=3.11.5',
    include_package_data=True,
    data_files=[
        ('', ['run.py', 'LICENSE.txt', 'README.md']), 
        ('assets', ['assets/icon.ico', 'assets/icon.png', 'icon_github.png']),  
        ('config', ['config/config.py']), 
        ('functions', ['functions/functions.py']),  
        ('locales', ['locales/en.json', 'locales/es.json', 'languages.json']), 
        ('src', ['src/a_Importar_archivos_a_excel.py', 'src/b_Modificar_nombres.py', 'c_Desbloquear_excel.py', 'main_menu.py']),  
    ],
     entry_points={
        'console_scripts': [
            'Mass_files_renaming = run:main',  
        ],
    },
)