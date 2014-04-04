# git_diff_xlsx

1. Place the file `git_diff_xlsx` in a folder
2. Add the following line to the global `.gitconfig` file:

```
    [diff "zip"]
    binary = True
    textconv = python c:/path/to/git_diff_xlsx.py
```

3. Add the following line to the repository's `.gitattributes` file:
    `*.xlsx diff=zip`
4. Now, typing `git diff` at the prompt will produce text versions
of Excel `.xlsx` files 

See https://wiki.ucl.ac.uk/x/P7MpAg for more details.
