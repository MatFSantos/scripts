# Dependencies:

    - pandas
    - reportlab
    - python-dotenv
    - openpyxl

# How to use this scripts
    
### Execute this commands
- Make a environment file based on .env.example:
  ``cp .env.example .env``
- fill in the environment file with the variables.

- Make a development environment with python command:
  ``python3 -m venv venv``
- Active the development environment:
    - In linux:
        ``source venv/bin/activate``
- Install the dependencies with pip package manager:
    ``pip install pandas reportlab python-dotenv openpyxl``
- DONE.
- Now, alter the params in ``app.py`` to send yours emails
- Execute ``python3 app.py``

- **Remember if**: the worksheets must be in folder ``planilhas`` and the image template to your certificates must be in folder ``templates``

