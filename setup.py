from setuptools import setup, find_packages

setup(
    name="olaqin-template",
    version="1.0.0",
    packages=find_packages(),
    install_requires=[
        "python-pptx==0.6.22",
    ],
    author="Olaqin",
    description="Template PowerPoint standardisé pour les présentations Olaqin",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    include_package_data=True,
    package_data={
        "": ["assets/*.png"],
    },
) 