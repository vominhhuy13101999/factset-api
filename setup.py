import setuptools 
setuptools.setup(
    name='FactSet_api',
    version='0.0.1',
    # packages=['API'],
    # description='build dcf model',
    # author='Huy Vo',
    # author_email='huy.vo@student.fairfield.edu',
    packages=setuptools.find_packages(),
    install_requires=[
        # 'wheel',
        'Xlsxwriter',
        'yfinance',
        
        
    ],
)