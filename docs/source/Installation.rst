.. _Installationer:

****************************************
Installation
****************************************

To use the SmkTool it is necessary to have Python installed.

.. note::
   You have to activate the virtual environment once for every workspace in vscode.


.. note::
   Currently we have an issue with the conan plugin!! Please install not a python version > 3.11!!


Installation for User
---------------------------------------


The first Step is to create an folder and join them:

.. code-block:: bash

    $ mkdir MyProjectTest
    $ cd MyProjectTest


The next Step is to install the virtual environment:

.. code-block:: bash

    $ pip install virtualenv
    $ virtualenv .venv

Then we have to activate the virtual environment:

.. code-block:: bash

    $ .venv\Scripts\activate.bat

The next step is to install the **SmkTool** via pip:

.. code-block:: bash

    $ pip install SmkTool

Generate a workspace

.. code-block:: bash

    $ SmkWorkspace


Now we can start vs code 

.. code-block:: bash

    $ code

To get the virtual environment running in vscode you have to select the python interpreter from the venv.
This can be done by pressing ctrl + shift + p. Then select Python: Select Interpreter.


Installation for Developer
---------------------------------------

Download the code from git

.. code-block:: bash

    $ git checkout https://git.marquardt.de/scm/aurixtc3x/smktool.git

The next step is to execute the prepare workspace script:

.. code-block:: bash

    $ PrepareWorkspace.bat

Now we can start vs code 

.. code-block:: bash

    $ code

To get the virtual environment running in vscode you have to select the python interpreter from the venv.
This ou can do by pressing ctrl + shift + p. Then select Python: Select Interpreter.