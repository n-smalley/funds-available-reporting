# Follow listed setup steps below:

## Create a local virtual environment

 $ py -m venv .venv (when VSCode asks to use the environment, "say Yes")

## Activate the virtual environment

  On Windows 
  
  $ .venv/Scripts/activate

## Install dependencies

  $ pip install -r requirements.txt --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org --default-timeout=1000
