import sys

# Make sure we are running python3.5+
if 10 * sys.version_info[0]  + sys.version_info[1] < 35:
    sys.exit("Sorry, only Python 3.5+ is supported.")

from setuptools import setup

def readme() -> str:
    with open('README.md') as f:
        return f.read()

if __name__ == "__main__":
    setup(
      long_description =   readme(),
      long_description_content_type = "text/markdown",
      include_package_data = True
    )
