"""
PyInstaller hook for numpy to avoid import errors
"""
from PyInstaller.utils.hooks import collect_submodules

# Only collect essential numpy submodules
hiddenimports = [
    'numpy.core._multiarray_umath',
    'numpy.core._multiarray_tests',
    'numpy._pyinstaller',
]

# Exclude unnecessary numpy modules
excludedimports = [
    'numpy.f2py',
    'numpy.distutils',
    'numpy.testing',
]
