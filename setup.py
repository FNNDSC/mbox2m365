import sys

# Make sure we are running python3.5+
if 10 * sys.version_info[0]  + sys.version_info[1] < 35:
    sys.exit("Sorry, only Python 3.5+ is supported.")

from setuptools import setup

def readme():
    with open('README.adoc') as f:
        return f.read()

setup(
      name             =   'mbox2m365',
      version          =   '1.0.0',
      description      =   'Send a message stored within an mbox using Office365',
      long_description =   readme(),
      author           =   'FNNDSC',
      author_email     =   'dev@babymri.org',
      url              =   'https://github.com/FNNDSC/mbox2m365',
      packages         =   ['mbox2m365', 'jobber'],
      install_requires =   ['pfmisc'],
      entry_points={
          'console_scripts': [
              'mbox2m365 = mbox2m365.__main__:main'
          ]
      },
      license          =   'MIT',
      zip_safe         =   False
)
