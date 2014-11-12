# http://stackoverflow.com/questions/17806485/execute-a-python-script-post-install-using-distutils-setuptools

from setuptools import setup, find_packages
setup(

	name = "gitxl",
	version = "0.1",
	packages = find_packages(exclude=['*test']),
	scripts = [],
	test_suite = "git_diff_xlsx.tests.test_all",

	install_requires = ['xlrd'],

	author = "Will Usher",
	author_email = "w.usher@ucl.ac.uk",
	description = "Converts Microsoft Excel files to text to enable easy \
	               git diff",
	license = "",
	keywords = "git excel",
	url = "https://github.com/willu47/git_diff_xlsx",





)
