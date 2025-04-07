from setuptools import setup, find_packages

title = "extractsql"
user = "cbaldelomar"
repo_url = f"https://github.com/{user}/{title}"

cli = f"{title}={title}.main:main"

requires = [
    "pyodbc>=5.1.0",
    "XlsxWriter>=3.2.0",
    "pyarrow>=18.1.0",
    "charset_normalizer>=3.4.1",
    "tqdm>=4.67.1",
]

encoding = "utf-8"

with open("README.md", encoding=encoding) as f:
    readme = f.read()

# Read the version number from __version__.py
version = {}
with open("extractsql/__version__.py", encoding=encoding) as f:
    exec(f.read(), version)

setup(
    name=title,
    version=version["__version__"],
    description="Extract data from SQL database into an output file.",
    long_description=readme,
    long_description_content_type="text/markdown",
    author=user,
    url=repo_url,  # GitHub repository URL
    packages=find_packages(),  # Automatically find modules
    python_requires=">=3.8",  # Specify minimum Python version
    install_requires=requires,
    entry_points={
        "console_scripts": [
            cli,  # CLI entry point
        ],
    },
    keywords="SQL Excel csv txt export data extraction",
    classifiers=[
        "Programming Language :: Python",
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)
