# docx2md

## Convert MS Word `.docx` files to Markdown

This is a simple enough *command-line* PHP script that will read a `.docx` file and convert it to Markdown format.

There are no tests written yet. This was put together rather quickly based on code snippets from a closed-source project I had written in 2011. It is also pretty trivial to extend this to work on a webserver. I went the easy route with regards to command-line only support so I wouldn't have to deal with managing file uploads.

Also, since most of the conversion work is left to XSL transformations, this utility could be very easily ported to another language other than PHP. I chose PHP because my existing code snippets were in PHP.

I look forward to any pull requests or bug reports to help make this converter as reliable as possible.