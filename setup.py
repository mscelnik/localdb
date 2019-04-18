""" Setup script for the LocalDB interface module.
"""

from setuptools import setup

version = '0.1'

setup(
    name='localdb',
    version=version,
    author='Matthew Celnik',
    author_email='matthew@celnik.co.uk',
    url='https://github.com/mscelnik/localdb',
    short_description='MSSQL SQLLocalDB.exe wrapper (Windows only)',
    description='Python wrapper of sqllocaldb.exe command line program.',
    license='MIT',
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Intended Audience :: Developers',
        'Topic :: Software Development :: Libraries',
        'License :: OSI Approved :: MIT License'
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
    ],
    keywords='mssql localdb',
    packages=[],
    py_modules=['localdb'],
    install_requires=[],
    requires=['pyodbc'],
    obsoletes=[f'localdb(<={version})'],
    zip_safe=False,  # Required for conda build to work on Python 3.7.x.
)
