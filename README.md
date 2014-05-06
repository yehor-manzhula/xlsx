xlsx
====

Xlsx generator for nodejs

## Contents: ##

- [Features](#Features)
- [Installation](#Installation)
- [Examples](#Examples)
- [Todo](#ToDo)
- [License](#Licence)
- [Credit](#Credit)

<a name="Features"/>
## Features: ##

- Generating Microsoft Excel document (.xlsx file):
  - Create Excel document with one or more sheets. Supporting cells of type both number and string.
  - Set bold text style

<a name="Installation"/>
## Installation: ##

via Git:

```bash
$ git clone git@github.com:egor-manjula/xlsx.git
```

This module is depending on:

- archiver
- mustache
- moment
- when

<a name="Examples"/>
## Examples: ##
- examples/generate_xlsx.js - Example how to create Excel 2007 sheet and save it into file.

<a name="Todo"/>
## Features todo: ##
- Implement other text styles

<a name="Licence"/>
## License: ##

(The MIT License)

Copyright (c) 2014 Manjula Egor;

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
'Software'), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

<a name="Credit"/>
## Credit: ##
- Inspired with Officegen nodejs library https://github.com/vtloc/officegen