import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="pyvba",
    version="0.6.0",
    author="Jack Wilson",
    description="Data mine and write scrips for VBA applications on Windows.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/TheEric960/pyvba",
    packages=setuptools.find_packages(),
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Environment :: Win32 (MS Windows)",
        "License :: OSI Approved :: MIT License",
        "Natural Language :: English",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3.7",
    ],
    install_requires=['pywin32'],
    python_requires='>=3.7',
)
