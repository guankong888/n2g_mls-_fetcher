# setup.py
from setuptools import setup

setup(
    name="n2g_mls_fetcher",
    version="0.1.0",
    description="Fetch N2G Master Location Sheet tabs via Microsoft Graph into pandas",
    py_modules=["mlsfetcher"],
    install_requires=[
        "requests>=2.25",
        "pandas>=1.3",
        "msal>=1.10"
    ],
    entry_points={
        "console_scripts": [
            "mlsfetcher=mlsfetcher:main",  # if you have a main() in mlsfetcher
        ]
    },
)
