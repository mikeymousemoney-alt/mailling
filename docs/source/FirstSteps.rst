.. _FirstSteps:

First Steps
================================================================================

If the SmkTool is installed as described in the installation chapter, the following steps can be carried out to generate a project.

Create the workspace
****************************

There are two ways to generate the project workspace. The simple way is to use the task from the SmkTool worcspace. 


With the help of the SmkTool workspace
------------------------------------------

Execute the task Generate Worcspace by press ctrl + shift + p followed by task and then select the generation task.
Now we can open the project workspace and select the venv in vs code.


On the command line
------------------------------------------

Create a venv and install the SmkTool.
Select the venv via:

.. code-block:: bash

    $ .venv\Scripts\activate.bat

Then execute the following command:

.. code-block:: bash

    $ SmkTool --project-name myProject --from-example bswsmk/examples/Cvm_HSM/SmkInput.json 

PS. don't forget the environment setup in vscode.


Setup the project workspace in vs code
************************************************

To get the virtual environment running in vscode you have to select the python interpreter from the venv.
This ou can do by pressing ctrl + shift + p. Then select Python: Select Interpreter.
