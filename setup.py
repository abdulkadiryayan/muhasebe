from setuptools import setup, find_packages

setup(
    name="muhasebe_app",
    version="0.1",
    packages=find_packages(),
    install_requires=[
        'tkcalendar',
        'openpyxl',
        'pandas',
    ],
) 