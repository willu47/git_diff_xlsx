# git_diff_xlsx

[![Build Status](https://travis-ci.org/willu47/git_diff_xlsx.svg?branch=develop)](https://travis-ci.org/willu47/git_diff_xlsx)

1. Run `python setup.py install`
2. Add the following line to the global .gitconfig file:

```
    [diff "zip"]
    binary = True
    textconv = python -m parse_xml
    cachetextconv = true
```

3. Add the following line to the repository's .gitattributes
    `*.xlsx diff=zip`
4. Now, typing `git diff` at the prompt will produce differences between
text versions of Excel `.xlsx` files
