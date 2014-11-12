from nose.tools import assert_equal, raises
from ..git_diff_xlsx import parse
import sys


def test_parse_xlsx():
    infile = "tests/test1.xlsx"
    outfile = sys.stdout
    actual = parse(infile, outfile)
    desired = ""
    assert_equal(actual, desired)
