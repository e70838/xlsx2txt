# xlsx2txt Tiny Perl script to convert xlsx files to text usable for a diff.

Hello, I have a couple of big excel files that are stored in version control system where I have many branches.

I have not found any good tool to find the differences between two excel files. I have wrote a small Perl script that produces a textual representation of the files that can be send to diff command.

This script is a simple use of module Spreadsheet::XLSX.

It is not installed by default on ubuntu, you can install the required modules using:
apt-get install libspreadsheet-xlsx-perl libtext-iconv-perl

All suggestions are welcome.

License: do what you want of this code.