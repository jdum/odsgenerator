from os.path import join, dirname, abspath
from setuptools import setup, find_packages


def content_of(rel_path):
    with open(join(abspath(dirname(__file__)), rel_path)) as f:
        return f.read()


def get_version(rel_path):
    for line in content_of(rel_path).splitlines():
        if line.startswith("__version__"):
            delim = '"' if '"' in line else "'"
            return line.split(delim)[1]
    raise RuntimeError("Unable to find version string.")


long_description = content_of("README.rst")


setup(
    name="odsgenerator",
    version=get_version("odsgenerator/odsgenerator.py"),
    description="Generate an OpenDocument Format .ods file from json or yaml file",
    long_description=long_description,
    long_description_content_type="text/x-rst",
    url="https://github.com/jdum/odsgenerator",
    author="Jérôme Dumonteil",
    author_email="jerome.dumonteil@gmail.com",
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Environment :: Console",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Operating System :: OS Independent",
        "Topic :: Text Processing",
    ],
    keywords="text json openDocument ODF ods generator",
    license="MIT",
    python_requires=">=3.6",
    packages=find_packages(),
    include_package_data=True,
    install_requires=["odfdo>=3.3", "pyyaml"],
    entry_points={"console_scripts": ["odsgenerator=odsgenerator.command_line:main"]},
)
