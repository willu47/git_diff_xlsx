# git_diff_xlsx  [![Build Status](https://travis-ci.org/willu47/git_diff_xlsx.svg?branch=develop)](https://travis-ci.org/willu47/git_diff_xlsx)


This script parses an .xlsx file and converts it into a text format which can be compared using git.

I wrote this script as I wanted to use a version control package for managing an existing computational model,
the input files of which are defined using Microsoft Excel workbooks.

See my [blog entry](https://wiki.ucl.ac.uk/x/P7MpAg) for more details of how it works.

## Installation

1. Download the [latest release](https://github.com/willu47/git_diff_xlsx/releases/latest)
1. Run `python setup.py install`
2. Add the following lines to the global .gitconfig file:

```
    [diff "git_diff_xlsx"]
    binary = True
    textconv = parse_xlsx
    cachetextconv = true
```

3. Add the following line to your repository's `.gitattributes` file
    `*.xlsx diff=git_diff_xlsx`
4. Now, typing `git diff` at the prompt will produce differences between
text versions of Excel `.xlsx` files

## Caveats

There are a bunch of issues with this script.
I wrote it to fulfil a need I had then and there and there are lots of hard-coded horrors.
Please feel free to contribute to cleaning up the code, submit issues and pull-requests.
