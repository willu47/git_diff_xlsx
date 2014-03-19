* git_diff_xlsx
1. Place the file `parse_xml.py` in a folder
2. Add the following line to the global .gitconfig file:
    [diff "zip"]
    binary = True
    textconv = python c:/path/to/parse_xml.py
3. Add the following line to the repository's .gitattributes
    *.xlsx diff=zip
4. Now, typing `git diff` at the prompt will produce text versions
of Excel .xlsx files