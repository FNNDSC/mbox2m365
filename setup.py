import sys

# Make sure we are running python3.5+
if 10 * sys.version_info[0]  + sys.version_info[1] < 35:
    sys.exit("Sorry, only Python 3.5+ is supported.")

from setuptools import setup

def readme():
    with open('README.adoc') as f:
        return f.read()

if __name__ == "__main__":
    setup(
      long_description =   readme(),
    )