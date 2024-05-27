from PyInstaller.utils.hooks import copy_metadata

# Ensure that the required plugins are copied
datas = copy_metadata('pyexcel-io')