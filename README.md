# docx2md

## Convert MS Word `.docx` files to Markdown

This is a simple enough *command-line* PHP script that will read a `.docx` file and convert it to Markdown format.

There are no tests written yet. This was put together rather quickly based on code snippets from a closed-source project I had written in 2011. It is also pretty trivial to extend this to work on a webserver. I went the easy route with regards to command-line only support so I wouldn't have to deal with managing file uploads.

Also, since most of the conversion work is left to XSL transformations, this utility could be very easily ported to another language other than PHP. I chose PHP because my existing code snippets were in PHP.

I look forward to any pull requests or bug reports to help make this converter as reliable as possible.

## License

Copyright (c) 2012 Mathieu Bouchard

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.