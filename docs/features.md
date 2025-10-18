# Features

`html2pptx` supports a wide range of HTML elements and CSS properties to create high-quality PowerPoint presentations.

## Supported HTML Elements

-   `<div>`, `<p>`, `<span>`
-   `<h1>` - `<h6>`
-   `<li>`, `<ul>`, `<ol>`
-   `<a>` (with `href` for hyperlinks)
-   `<img>`
-   `<svg>` (including `<line>` and `<text>`)

## Supported CSS Properties

### Layout
-   `display: flex`
-   `display: grid`
-   `flex-direction`
-   `grid-template-columns`
-   `grid-template-rows`
-   `gap`, `column-gap`, `row-gap`
-   `position` (absolute, relative, fixed)
-   `top`, `left`, `right`, `bottom`

### Box Model
-   `width`, `height`
-   `padding`
-   `border`, `border-color`, `border-width`, `border-style`
-   `border-radius`
-   `background-color`, `background`

### Text Styling
-   `font-size`, `font-family`, `font-weight`, `font-style`
-   `color`
-   `text-align`
-   `align-items` (for flexbox)
-   `justify-content` (for flexbox)

### Transforms
-   `transform: rotate(...)`
-   `transform: scale(...)`

### Pseudo-selectors
-   `:nth-child(n)`
-   `:first-child`
-   `:last-child`