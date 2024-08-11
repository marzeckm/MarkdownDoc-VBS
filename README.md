# MarkdownDoc (VBS)
![VBScript](https://img.shields.io/badge/vbscript-black?style=for-the-badge&logo=data:image/svg%2bxml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz4KPHN2ZyB2aWV3Qm94PSIwIDAgNDQ4IDUxMiIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbD0iI2ZlZmVmZSIgZD0iTTAgMzJoMjE0LjZ2MjE0LjZIMFYzMnptMjMzLjQgMEg0NDh2MjE0LjZIMjMzLjRWMzJ6TTAgMjY1LjRoMjE0LjZWNDgwSDBWMjY1LjR6bTIzMy40IDBINDQ4VjQ4MEgyMzMuNFYyNjUuNHoiLz4KPC9zdmc+Cg==)
  
MarkdownDoc (VBS) is a lightweight tool designed to simplify the creation of documentation for software or other projects using Markdown. This small yet powerful script, just under 20 KB, effortlessly converts your Markdown files into fully formatted HTML documentation that can be hosted on your website. The script is highly compatible with almost all Windows versions since Windows XP.
  
## Getting Started
To start using MarkdownDoc, simply download the code by clicking on the `Code` button and selecting `Download as Zip`. After downloading, extract the folder using WinZip, 7-Zip, WinRAR, or any other archiving tool. Navigate to the extracted folder, open the `src` directory, and locate the `index.md` file. You can open `index.md` with a text editor such as Notepad, Notepad++, or Visual Studio Code to begin creating your documentation in Markdown.
  
## Supported Features
MarkdownDoc (VBS) supports the core features of Markdown, allowing you to create comprehensive documentation with ease.
  
### Headers
You can create headers from H1 to H6 using the standard Markdown syntax:
  
```
# This is a Header of Type H1

## This is a Header of Type  H2

### This is a Header of Type  H3

#### This is a Header of Type  H4

##### This is a Header of Type H5

###### This is a Header of Type H6
```
  
### Paragraphs
No special syntax is required for regular text. Simply type your text, and it will be formatted into paragraphs. To insert a line break, use two spaces at the end of the line.
  
```
Lorem ipsum dolor sit amet consectetur adipisicing elit. Debitis officia distinctio cupiditate aperiam earum ullam a dicta illo ad modi obcaecati adipisci exercitationem deleniti doloremque vitae, eum esse impedit eligendi. Amet ullam voluptatibus repudiandae minima eum laborum enim? Suscipit incidunt nostrum porro autem sapiente voluptatum vitae doloribus mollitia, amet saepe.  
Lorem ipsum dolor sit amet consectetur adipisicing elit. Debitis officia distinctio cupiditate aperiam earum ullam a dicta illo ad modi obcaecati adipisci exercitationem deleniti doloremque vitae.
```
  
### Bold Text
To make text bold, surround the text with two asterisks (`**`). This will be converted into <strong> tags in HTML.
  
```
This is normal text, but **here the text is bold**.
```
  
### Italic Text
Similarly, to italicize text, use one asterisk (`*`) on each side of the text. This will be converted into <em> tags in HTML.
  
```
This is normal text, but *here the text is italic*.
```
  
### Italic & Bold Text
For text that is both bold and italic, use three asterisks (`***`) on each side.
  
```
This is normal text, but ***here the text is bold/italic***.
```
  
### Blockquotes
To create blockquotes, start the line with a greater-than symbol (`>`). Nested blockquotes are currently not supported.
 
```
> This is a blockquote text
```
  
### Lists
MarkdownDoc supports both ordered and unordered lists.
  
#### Ordered Lists
Create an ordered list by starting lines with numbers followed by a period.
  
```
1. This is the first item
2. This is the second item
3. This is the third item
```
  
#### Unordered Lists
To create an unordered list, use `*`, `+` or `-` at the beginning of each line. Be cautious when using `*` to avoid conflicts with bold text syntax.
  
```
- This is a list item
- This is a list item
  
* This is a list item
* This is a list item
  
+ This is a list item
+ This is a list item
```
  
### Links
You can include hyperlinks using the standard Markdown syntax. Links open in a new tab by default.
    
```
[Hyperlink-Text](https://github.com/marzeckm)
```
  
### Images
You can insert images in Markdown either by linking to web images or by adding images to the `src` folder and linking them relatively.
  
```
This is an image from the web  
![Alt-Text](https://th.bing.com/th/id/OIP.6bWOQzWJ8PAdd-f70-oRbAHaEK?rs=1&pid=ImgDetMain)  
  
This is a local image
![Alt-Text](./src/pexels-markusspiske-2004161.jpg)
```
  
## Usage Instructions
To build your documentation, simply create or edit the `index.md` file in the `/src` folder. You can add images, PDFs, or other files that you want to link in your documentation. Once you’re ready to generate the HTML, run the `generator.vbs` script. On newer Windows versions like Windows 7 and later, you can execute the script by double-clicking it to open a terminal window with the program running.

Alternatively, you can manually open CMD or PowerShell, navigate to the script’s directory, and run it. Here's how to do it on Windows 7, 8, 8.1, 10, and 11:
  
```
cd "C://Path/To/The/Generator/"
./generator.vbs
```
  
Make sure to replace `"C://Path/To/The/Generator/"` with the actual path to your script.
  
For older versions like Windows 98, ME, 2000, XP, and Vista, you may need to run the script using `Cscript`. While the script will work without `Cscript`, you won’t see the program’s output. To run it with `Cscript`, use the following command:
  
```
cd "C://Path/To/The/Generator/"
CScript generator.vbs
```
  
### Open the Documentation
After generating the HTML documentation, you can open it in a web browser that is at least as modern as Internet Explorer 11.
  
#### Windows 10, Windows 11
On Windows 10 and 11, you can use the pre-installed Microsoft Edge or other browsers like Google Chrome, Chromium, or Mozilla Firefox.
    
#### Windows 7, Windows 8, Windows 8.1
For Windows 7, 8, and 8.1, you can use Internet Explorer 11 or supported versions of Microsoft Edge, Google Chrome, Chromium, Mozilla Firefox, or Supermium.
  
#### Windows Vista, Windows XP
On Windows Vista or XP, you can use the Supermium (based on Chromium) or MyPal (based on Firefox) browsers, which support these older systems.
  
#### Windows 2000, Windows ME, Windows 98
To view the documentation on a Windows 2000 machine, you can use K-Meleon 74. For instructions on installing it, [see this guide](http://kmeleonbrowser.org/wiki/InstallerForWindows2000). On Windows ME and 98, you can use K-Meleon 1.5.4 without kernel extension. [Here's how to install it](http://kmeleonbrowser.org/wiki/InstallerForWindows98).
  
## Requirements
- Text-Editor for editing the code files
- Windows 98, Windows NT4, Windows ME, Windows 2000, Windows XP, Windows Server 2003, Windows Vista, Windows Server 2008 (R2), Windows 7, Windows 8/8.1, Windows Server 2012 (R2), Windows 10, Windows Server 2016, Windows Server 2019, Windows 11, Windows Server 2022 or Windows Server 2025
- Microsoft Edge, Google Chrome, Mozilla Firefox, Internet Explorer 11, Chromium, Supermium, K-Meleon
  
## Contribute
If you want to contribute to the development of this project, feel free to submit pull requests or open issues. Let's make the MarkdownDoc (VBS) even better together!
  
## License
This project is licensed under the [MIT License](LICENSE).