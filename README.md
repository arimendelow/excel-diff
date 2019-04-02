# excel-diff by Ari Mendelow
### Inspired by [this project](https://matthewkudija.com/blog/2018/07/21/excel-diff/)

Very simple to use, but requires a level of comfortability in command line. More thorough instructions with screenshots to come soon, but for now, a description of the files included in this repository, and a few basic instructions:

- `excel-diff.py` - the source code for this script. Requires various imports to be run (see the `imports` section), and relies on Python3.
- `excel-diff.exe` - a standalone (requires Python installation) Windows executable of the above script. Must be run from command line. Was created using [Nuitka](https://github.com/Nuitka/Nuitka).

Both of the above files are run the same way:
- `python excel-diff.py <file1.xlxs> <file2.xlxs>`
- `excel-diff.exe <file1.xlxs> <file2.xlxs>`

A file of the format `<file1.xlxs> vs <file2.xlxs>.xlxs` will be created in the present working directory (pwd) of the terminal/command prompt from which the script/executable is called. Therefore, it is convenient to include the script/executable in the same file as the excel files being compared.

Thats all for now; a more thorough walkthrough is currently in development. Enjoy!
