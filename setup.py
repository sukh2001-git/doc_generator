from setuptools import setup, find_packages

with open("requirements.txt") as f:
    install_requires = f.read().strip().split("\n")

# get version from __version__ variable in doc_generator/__init__.py
from doc_generator import __version__ as version

setup(
    name="doc_generator",
    version=version,
    description="Doc Generator",
    author="OneHash Technologies Pvt. Ltd.",
    author_email="sukhman@onehash.ai",
    packages=find_packages(),
    zip_safe=False,
    include_package_data=True,
    install_requires=install_requires,
)