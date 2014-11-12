from nose.tools import assert_equal
from git_diff_xlsx import main

def test_parse_xlsx():
    actual = parse(infile, outfile)
    desired =
    assert_equal(actual, desired)
