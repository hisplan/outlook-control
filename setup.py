import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

# get __version__
exec(open("src/outlook/version.py").read())

setuptools.setup(
    name="outlook-control",
    version=__version__,
    author="Jaeyoung Chun",
    author_email="chunj@mskcc.org",
    description="outlook-control",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/hisplan/outlook-control",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.6",
    # packages=setuptools.find_packages(),
    packages=["outlook"],
    package_dir={"": "src"},
    install_requires=["appscript==1.1.1"],
)
