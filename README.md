# Open and Save Archive
(_formerly **xslt for docx**_)

Joel Kalvesmaki, 2019 ([licensing information](LICENSE.md) | [changes](CHANGES.md))

XSLT to handle **docx** (Microsoft Word), **xlsx** (Microsoft Excel), **zip**, **epub**, **odt**, **ods**, **jar**, **rar**, and **all sorts of compressed archive formats**.

- [Saxon PE and EE](http://saxonica.com/products/products.xml): any kind of archive can be retrieved or saved.
- [Saxon HE](http://saxonica.com/products/products.xml): only docx and xlsx formats can be opened and saved, and without any binary components (images, videos, etc.).

The only stylesheet you need is [open-and-save-archive.xsl](open-and-save-archive.xsl). The others, [open-and-save-xlsx.xsl](open-and-save-xlsx.xsl) and [open-and-save-docx.xsl](open-and-save-docx.xsl), are mere aliases. 

No Ant. No XProc. Open and Save Archive is simple, pure XSLT (3.0), tapping the power of [EXPath](http://expath.org).

## Examples

Practical applications are featured in the example subdirectories, described here:

1. [Unpacking and saving archives](example%201,%20unpacking%20and%20saving%20various%20archives): basic demonstration of how to fetch the component parts of an archive, then to repackage and save them. This demonstration shows the variety of archive types that can be handled. The remaining examples (##2-5 below) deal only with docx and xlsx files. 
1. [Plain text](example%202%2C%20get%20plain%20text): shows how to scrape multiple docx or xlsx files for their plain text content and concatenate it in a single file.
1. [Replacing text via regular expressions](example%203%2C%20change%20with%20regular%20expressions): shows how to search and replace within a Word or Excel file using regular expressions. This example is important because regular expressions are non-existent in Excel, and deficient in Word. Finding and replacing text in the XML components of a docx file is tricky, and I use and explain in this example what I call the Splintered Seas Technique (with apologies to anyone who might have invented, used, and named a similar technique before me).  
1. [Make form letters](example%204%2C%20make%20form%20letters): shows how to turn a template Word document and an XML database into form letters. This example is important because Word cannot easily handle data that does not fit the spreadsheet model, and does not have good tools for coordinating and manipulating data. In this example, you use XSLT to define variables of your choice, then you type those variable names wherever you like within the docx template, e.g., *$family-name*. You can iterate over multiple values, and apply XSLT functions to change the data and its formatting as you like--things that are difficult or impossible to do in Word.
1. [Anonymize documents](example%205,%20anonymize%20document): shows how to quickly scrub from the metadata of a document the names of those who are credited with changes, comments, or tracked changes. This is useful when returning to an author a manuscript that has been through blind peer review, and you wish to preserve the anonymity of the writers. To my knowledge this functionality is missing from Word.

Under the hood, Word and Excel files can get quite complex. The XSLT files in the examples have been written specifically for the accompanying sample input. You may need to develop the code to handle the input you are working with.  

## Notes

_(See the [main stylesheet](open-and-save-archive.xsl) for further notes.)_ 

Functions are in the [Text Alignment Network (TAN)](http://textalign.net) namespace, `tag:textalign.net,2015:ns`.

Saxon HE does not support extended functions, such as reading and writing binary files, or getting the contents of an archive. To retrieve all the components of an archive with Saxon HE, one must know how how it is structured. I have written routines that will get everything from docx and xslx architecture, but other file types could be handled. You are welcome to extend the stylesheets to handle the archive types you like.

Opening an archive (`tan:open-archive()`) returns its components as a sequence of XML documents. If a component is binary, it will be returned as a text node of the root element in base 64 binary. Each root element has temporary attributes `@xml:base` pointing to the resolved uri of the archive and `@_archive-path` pointing to the relative place of the component.

With Saxon PE and EE, you can open the archive as a map, instead of a sequence of documents, if you like. See examples featuring `tan:extract-map($example-archives[1], $example-archive-entries-maps[1])`. 

Saving an archive (`tan:save-archive()`) requires as input the archive components as a sequence of XML documents, each with an `@_archive-path` in the root element to indicate where in the archive the component should be placed. Currently the saved output is a single archive, with the components of originally nested archives themselves not archived. Reconstructing nested archives may be supported in a future release.

`tan:extract-map()` is my attempt to instantiate, enhance, and develop the EXPath function `arch:extract-map()`. See the stylesheet for more comments. 

Do you find working with maps in XSLT bothersome? You may find the companion function `tan:map-to-xml()`, which lets you convert a map to an XML tree, to be extremely useful.

You can safely either include or import the key [stylesheet](open-and-save-archive.xsl).  It does not declare or define an initial template or default template behavior, so it shouldn't interfere with any stylesheet that includes or imports it. But you will need to make sure that the including/importing stylesheet does not itself interfere with Open and Save Archive:

If you *include* it (the equivalent of copying the code directly in the including XSLT), watch out for how the default template behavior is defined in the including module, because there may be template-rule conflicts. Watch out for values, default or explicitly declared, of `@priority`, and the behavior of `@mode='#all'`.

If you *import* it (a softer form of inclusion, where rules and parameters specified in the imported file can easily be ignored or overwritten), be certain to add something like the following:

    <xsl:template match="document-node() | node() | @*" priority="2"
        mode="clean-up-archive map-to-xml">
        <xsl:apply-imports/>
    </xsl:template>
    
In the above code, you might be able to dispense with `@priority`, or you might need to change its value. It depends upon what's happening in your master XSLT file.  
 