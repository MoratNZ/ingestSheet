
try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

config = {
    'description': 'My Project',
    'author': 'David Maclagan',
    'url': 'URL to get it at.',
    'download_url': 'Where to download it.',
    'author_email': 'david@verso.org.',
    'version': '0.1',
    'install_requires': ['openpyxl'],
    'packages': ['ingestSheet'],
    'scripts': [],
    'name': 'ingestSheet'
}

setup(**config)
