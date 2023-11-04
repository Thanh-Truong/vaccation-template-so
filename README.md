This is sample package
# Run the code
````
$ python foo/bar/main.py 
Hello World
````

# Install and run it locally

Setup
````
# Setup
$ python setup.py install

# Run
$ samplepackage

# It should output
Hello World
````

# Packaging and uploading

Preparation
````
$ python3 -m venv env
$ source env/bin/activate
$ pip install --upgrade pip build twine

````
then following the instruction here to upload the built package to PyPI or TestPyPI
https://thanh-truong.github.io/posts/2021/06/package-your-python-code/