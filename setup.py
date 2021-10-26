import setuptools

setuptools.setup(
    name='excel_extractor',
    version='1.0',
    author='Saba Tazayoni',
    author_email="",
    description="A package to create relational databases via Microsoft Excel.",
    url="https://github.com/stazay/excel_extractor",
    packages=setuptools.find_packages(),
    install_requires=['datetime',
                      'xlsxwriter',
                      'xlwings'], 
    python_requires='>=3.6',
    ) 
