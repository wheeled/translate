#!/usr/bin/env python

"""Setup script for packaging translate.

To build a package for distribution:
    python setup.py sdist
and upload it to the PyPI with:
    python setup.py upload

Install a link for development work:
    pip install -e .

Thee manifest.in file is used for data files.

"""

import os

from setuptools import setup, find_packages

here = os.path.abspath(os.path.dirname(__file__))
try:
    with open(os.path.join(here, 'README.rst')) as f:
        README = f.read()
except IOError:
    README = ''

try:
    from importlib.util import module_from_spec, spec_from_file_location
    spec = spec_from_file_location("constants", "./translate/_constants.py")
    constants = module_from_spec(spec)
    spec.loader.exec_module(constants)
except ImportError:
    # python2.7
    import imp
    constants = imp.load_source("constants", "./translate/_constants.py")

__author__ = constants.__author__
__author_email__ = constants.__author_email__
__license__ = constants.__license__
__maintainer_email__ = constants.__maintainer_email__
__url__ = constants.__url__
__version__ = constants.__version__


setup(name='translate',
    packages=find_packages(),
    # metadata
    version=__version__,
    description="A Python utility to use Google Cloud Translate API to translate documents from one language to another",
    long_description=README,
    author=__author__,
    author_email=__author_email__,
    url=__url__,
    license=__license__,
    python_requires=">=2.7, !=3.0.*, !=3.1.*, !=3.2.*, !=3.3.*, !=3.4.*",  # need to verify this ...
    # install_requires=[
    #     'jdcal', 'et_xmlfile',
    #     ],
    classifiers=[
                 'Development Status :: 3 - Alpha',
                 'Operating System :: MacOS :: MacOS X',
                 'Operating System :: Microsoft :: Windows',
                 'Operating System :: POSIX',
                 'License :: OSI Approved :: MIT License',
                 'License :: OSI Approved :: Apache License',
                 'Programming Language :: Python',
                 'Programming Language :: Python :: 3.6',
                 'Programming Language :: Python :: 3.7',
                 ],
    )
