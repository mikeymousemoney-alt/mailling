@REM pip install -e .
@REM del docs\source\generated\*.rst
rmdir doc /s /q
@REM sphinx-autogen -o docs\source\generated docs\source\api.rst
sphinx-build -b html docs\source doc 