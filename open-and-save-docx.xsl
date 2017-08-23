<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet exclude-result-prefixes="#all" xmlns:tan="tag:textalign.net,2015:ns"
    xmlns:html="http://www.w3.org/1999/xhtml" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
    xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:w10="urn:schemas-microsoft-com:office:word"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
    xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
    xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
    xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="2.0">

    <!-- Written 30 November 2016 by Joel Kalvesmaki, released under a Creative Commons 4.0SA license. -->

    <!-- These XSLT functions allow one to retrieve the XML documents that are part of a Word document and save component parts as a new Word document. -->
    
    <!-- All nontextual content is necessarily ignored, since XSLT deals only with texts. One work-around is to link to nontextual content, and not embed it. -->
    
    <!-- A key factor in making these functions successful is the introduction of @jar-path to the root element of every component XML document. The @jar-path indicates where in the hierarchy of the Word file each document sits. Only XML documents that retain @jar-path can be used by tan:save-docx(), and in those cases, the provisional @jar-path is removed before zipping. -->
    
    <!-- To see how these functions can be used, e.g., to alter a Word document using regular expressions, see the example subdirectory. -->

    <!-- OPENING WORD DOCUMENT FILES -->

    <xsl:function name="tan:docx-file-available" as="xs:boolean">
        <!-- Input: any element with an @href -->
        <!-- Output: a boolean indicating whether the Word document is available -->
        <xsl:param name="element-with-attr-href" as="element()?"/>
        <xsl:variable name="input-base-uri" select="base-uri($element-with-attr-href)"/>
        <xsl:variable name="static-base-uri" select="static-base-uri()"/>
        <xsl:variable name="best-uri" select="($input-base-uri, $static-base-uri)[1]"/>
        <xsl:variable name="this-href" select="$element-with-attr-href/@href"/>
        <xsl:variable name="source-uri" select="resolve-uri($this-href, $best-uri)"/>
        <xsl:variable name="source-jar-uri" select="concat('zip:', $source-uri, '!/')"/>
        <xsl:variable name="source-root" select="concat($source-jar-uri, '_rels/.rels')"/>
        <xsl:copy-of select="doc-available($source-root)"/>
    </xsl:function>
    <xsl:function name="tan:open-docx" as="document-node()*">
        <!-- Input: any element with the attribute @href pointing to a Microsoft Office file -->
        <!-- Output: a sequence of the XML documents found inside the input (the main .rels file first, then the document .rels, then the source content types, then every file ending in .xml) To facilitate the reconstruction of the Word file, every extracted document will be stamped with @jar-path, with the local path and name of the component. -->
        <xsl:param name="element-with-attr-href" as="element()?"/>
        <xsl:variable name="input-base-uri" select="base-uri($element-with-attr-href)"/>
        <xsl:variable name="static-base-uri" select="static-base-uri()"/>
        <xsl:variable name="best-uri" select="($input-base-uri, $static-base-uri)[1]"/>
        <xsl:variable name="this-href" select="$element-with-attr-href/@href"/>
        <xsl:variable name="source-uri" select="resolve-uri($this-href, $best-uri)"/>
        <xsl:variable name="source-jar-uri" select="concat('zip:', $source-uri, '!/')"/>
        <xsl:variable name="source-root-rels-path" select="concat($source-jar-uri, '_rels/.rels')"/>
        <xsl:variable name="source-root-rels"
            select="tan:extract-docx-component($source-jar-uri, '_rels/.rels')"/>
        <xsl:variable name="source-word-rels"
            select="tan:extract-docx-component($source-jar-uri, 'word/_rels/document.xml.rels')"/>
        <xsl:variable name="other-possible-word-rel-names" as="xs:string*"
            select="('endnotes.xml.rels', 'footnotes.xml.rels', 'footer1.xml.rels', 'footer2.xml.rels', 'header1.xml.rels', 'header2.xml.rels', 'settings.xml.rels')"/>
        <xsl:variable name="source-word-rels-misc" as="document-node()*">
            <xsl:for-each select="$other-possible-word-rel-names">
                <xsl:copy-of
                    select="tan:extract-docx-component($source-jar-uri, concat('word/_rels/', .))"/>
            </xsl:for-each>
        </xsl:variable>
        <xsl:variable name="source-content-types"
            select="tan:extract-docx-component($source-jar-uri, '[Content_Types].xml')"/>
        <xsl:variable name="source-docs"
            select="
                (for $i in $source-root-rels//@Target[matches(., '\.xml$')]
                return
                    tan:extract-docx-component($source-jar-uri, $i),
                for $j in $source-word-rels//@Target[matches(., '\.xml$')]
                return
                    tan:extract-docx-component($source-jar-uri, concat('word/', $j)))"/>
        <xsl:choose>
            <xsl:when test="not(doc-available($source-root-rels-path))">
                <xsl:document>
                    <error>No document found at <xsl:value-of select="$source-jar-uri"/></error>
                </xsl:document>
            </xsl:when>
            <xsl:otherwise>
                <xsl:sequence
                    select="$source-root-rels, $source-word-rels, $source-word-rels-misc, $source-content-types, $source-docs"
                />
            </xsl:otherwise>
        </xsl:choose>
    </xsl:function>
    <xsl:function name="tan:extract-docx-component" as="document-node()?">
        <!-- Input: the base jar uri for a Word document; a path to a component part of a Word document -->
        <!-- Output: the XML document itself, but with @jar-path stamped into the root element -->
        <xsl:param name="source-jar-uri" as="xs:string"/>
        <xsl:param name="component-path" as="xs:string"/>
        <xsl:variable name="extracted-doc" as="document-node()?"
            select="
                if (doc-available(concat($source-jar-uri, $component-path))) then
                    doc(concat($source-jar-uri, $component-path))
                else
                    ()"/>
        <xsl:if test="exists($extracted-doc)">
            <xsl:document>
                <xsl:apply-templates select="$extracted-doc" mode="stamp-docx-component-with-path">
                    <xsl:with-param name="path" select="$component-path"/>
                </xsl:apply-templates>
            </xsl:document>
        </xsl:if>
    </xsl:function>
    <xsl:template match="comment() | processing-instruction()" mode="stamp-docx-component-with-path clean-up-word-file-before-repackaging">
        <xsl:copy-of select="."/>
    </xsl:template>
    <xsl:template match="*" mode="stamp-docx-component-with-path">
        <xsl:copy>
            <xsl:copy-of select="@*"/>
            <xsl:apply-templates mode="#current"/>
        </xsl:copy>
    </xsl:template>
    <xsl:template match="/*" mode="stamp-docx-component-with-path">
        <xsl:param name="path" as="xs:string"/>
        <xsl:copy>
            <xsl:copy-of select="@*"/>
            <xsl:attribute name="jar-path" select="$path"/>
            <xsl:apply-templates mode="#current"/>
            <!--<xsl:copy-of select="node()"/>-->
        </xsl:copy>
    </xsl:template>


    <!-- SAVING WORD DOCUMENT FILES -->

    <xsl:template name="tan:save-docx">
        <!-- Input: a sequence of documents that each have @jar-path stamped in the root element (the result of tan:open-docx()); a resolved uri for the new Word document -->
        <!-- Output: a file saved at the place located -->
        <!-- Ordinarily, this template would be a function, but <result-document> always fails in the context of a function. -->
        <xsl:param name="docx-parts" as="document-node()*"/>
        <xsl:param name="resolved-uri" as="xs:string"/>
        <xsl:for-each select="$docx-parts/*[@jar-path]">
            <xsl:result-document href="{concat('zip:', $resolved-uri, '!/', @jar-path)}">
                <xsl:document><xsl:apply-templates select="." mode="clean-up-word-file-before-repackaging"/></xsl:document>
            </xsl:result-document>
        </xsl:for-each>
    </xsl:template>

    <xsl:template match="*" mode="clean-up-word-file-before-repackaging">
        <xsl:copy>
            <xsl:copy-of select="@*"/>
            <xsl:apply-templates mode="#current"></xsl:apply-templates>
        </xsl:copy>
    </xsl:template>
    <xsl:template match="/*" mode="clean-up-word-file-before-repackaging">
        <!-- get rid of the special @jar-path we added, to automate repackaging in the right locations -->
        <xsl:copy>
            <xsl:copy-of select="@* except @jar-path"/>
            <!-- copying the attributes should also ensure that namespace nodes are copied; Word will mark a file as corrupt if otiose namespace nodes aren't included -->
            <xsl:apply-templates mode="clean-up-word-file-before-repackaging"/>
        </xsl:copy>
    </xsl:template>
    <xsl:template mode="clean-up-word-file-before-repackaging"
        match="rel:Relationship[root()/*/@jar-path = '_rels/.rels' and matches(@Target, '\.jpe?g$') and not(@TargetMode = 'External')]">
        <!-- get rid of elements in .rels that point to images (e.g., thumbnails) -->
    </xsl:template>
    
</xsl:stylesheet>
