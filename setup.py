from setuptools import setup, find_packages

setup(
    name="word-to-latex-converter",
    version="1.0.0",
    packages=find_packages(),
    install_requires=[
        "pypandoc>=1.13",
    ],
    entry_points={
        "console_scripts": [
            "word2latex=main:run_app",
        ],
    },
)