# API Reference

This section provides a detailed reference for the `html2pptx` library's API.

## `convertHTML2PPTX(inputPath, outputPath, options)`

A convenience function for quick conversions.

-   **`inputPath`** (string): The path to the input HTML file.
-   **`outputPath`** (string): The path where the output PPTX file will be saved.
-   **`options`** (object, optional): An options object to customize the conversion. See below for details.

## `new HTML2PPTX(options)`

Creates a new instance of the `HTML2PPTX` converter.

-   **`options`** (object, optional): An options object.

### Options

-   **`slideWidth`** (number): The width of the PowerPoint slides in inches. Default: `10`.
-   **`slideHeight`** (number): The height of the PowerPoint slides in inches. Default: `5.625`.
-   **`htmlWidth`** (number): The width of the HTML container in pixels, used for scaling. Default: `1280`.
-   **`htmlHeight`** (number): The height of the HTML container in pixels, used for scaling. Default: `720`.

## `converter.convert(inputPath, outputPath)`

Converts an HTML file using the settings from the `HTML2PPTX` instance.

-   **`inputPath`** (string): The path to the input HTML file.
-   **`outputPath`** (string): The path for the output PPTX file.