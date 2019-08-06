from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
buildOptions = dict(packages = [], excludes = [], include_files = ["myicon.ico", 'doc', 'tools'])

import sys
base = 'Win32GUI' if sys.platform=='win32' else None

executables = [
    Executable('office2picture.py', base=base, compress = True, icon = "myicon.ico")
]
# Executable('newimages.py', base=base, compress = True, icon = "newimages.ico"), Executable('update.py', base=base, compress = True)
# Executable('newimages.py', base=base,compress = True, icon = "newimages.ico")
setup(name='office2picture',
      version = '1.0',
      description = 'Convert Office document to picture tools',
      author = 'Face2group.com',
      author_email = 'luckrill@163.com',	  
      options = dict(build_exe = buildOptions),
      executables = executables)
