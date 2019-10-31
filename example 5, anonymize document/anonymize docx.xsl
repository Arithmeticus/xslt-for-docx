<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:arch="http://expath.org/ns/archive" 
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:file="http://expath.org/ns/file"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:tan="tag:textalign.net,2015:ns"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:prop="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    xmlns:ssh="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    exclude-result-prefixes="#all" version="3.0">

    <xsl:param name="diagnostics-on" select="false()" static="true"/>

    <xsl:import href="../open-and-save-docx.xsl"/>
    <!-- The following inclusion has components common to stylesheets used for other examples. -->
    <xsl:include href="../inclusions/examples%20core.xsl"/>
    
    <!-- This stylesheet illustrates how to anonymize a Word file. -->
    <!-- Catalyzing input: any XML file (including this very one) -->
    <!-- Main input: the documents pointed to by $example-urls -->
    <!-- Primary output: none, unless diagnostics are on -->
    <!-- Secondary output: each file, with author and commenter information anonymized, saved in the output directory -->


    <!-- This application finds personal names in a document and replaces them with Author_ and a number -->

    <!-- This stylesheet has been developed solely for the accompanying example. Docx files can get quite 
        complicated, and your test case may require extra programming beyond the initial ideas outlined here. -->

    <xsl:output indent="yes" use-when="$diagnostics-on"/>
    
    <xsl:variable name="master-static-base-uri" select="static-base-uri()"/>

    <xsl:variable name="example-urls" as="element()?">
        <docx href="Lorem%20ipsum.docx"/>
    </xsl:variable>
    
    <xsl:variable name="authors-by-date" as="element()*">
        <xsl:for-each-group select="$example-components//(@*:author), $example-components/cp:coreProperties/(dc:creator, cp:lastModifiedBy)" group-by="(../@*:date, '9999')[1]">
            <xsl:sort select="current-grouping-key()"/>
            <xsl:variable name="this-date" select="current-grouping-key()"/>
            <xsl:for-each-group select="current-group()" group-by=".">
                <author date="{$this-date}">
                    <xsl:value-of select="current-grouping-key()"/>
                </author>
            </xsl:for-each-group> 
        </xsl:for-each-group> 
    </xsl:variable>
    <xsl:variable name="distinct-author-names" as="xs:string*">
        <xsl:for-each-group select="$authors-by-date" group-by=".">
            <xsl:value-of select="."/>
        </xsl:for-each-group> 
    </xsl:variable>
    
    <!-- STEP: ANONYMIZE THE METADATA -->
    <xsl:variable name="revised-components" as="document-node()*">
        <xsl:apply-templates select="$redirected-components" mode="anonymize-authors"/>
    </xsl:variable>
    
    <xsl:template match="@*:author" mode="anonymize-authors">
        <xsl:variable name="this-author-name" select="."/>
        <xsl:variable name="this-author-index" select="index-of($distinct-author-names, $this-author-name)"/>
        <xsl:attribute name="{name(.)}">
            <xsl:value-of select="'author_' || string($this-author-index)"/>
        </xsl:attribute>
    </xsl:template>
    <xsl:template match="@*:initials" mode="anonymize-authors">
        <xsl:variable name="this-author-name" select="(../@*:author)[1]"/>
        <xsl:variable name="this-author-index" select="index-of($distinct-author-names, $this-author-name)"/>
        <xsl:attribute name="{name(.)}">
            <xsl:value-of select="'A' || string($this-author-index)"/>
        </xsl:attribute>
    </xsl:template>
    <xsl:template match="dc:creator | cp:lastModifiedBy" mode="anonymize-authors">
        <xsl:variable name="this-author-name" select="."/>
        <xsl:variable name="this-author-index" select="index-of($distinct-author-names, $this-author-name)"/>
        <xsl:copy>
            <xsl:copy-of select="@*"/>
            <xsl:value-of select="'author_' || string($this-author-index)"/>
        </xsl:copy>
    </xsl:template>
    <xsl:template match="prop:Company" mode="anonymize-authors">
        <xsl:copy>
            <xsl:copy-of select="@*"/>
            <xsl:text> </xsl:text>
        </xsl:copy>
    </xsl:template>
    <!-- Eliminate any info in the people.xml component -->
    <xsl:template match="w15:person" mode="anonymize-authors"/>

    <!-- INITIAL TEMPLATE -->
    <xsl:template match="/">
        <xsl:choose>
            <xsl:when test="$diagnostics-on">
                <diagnostics>
                    <uris-resolved><xsl:value-of select="$example-uris-resolved"/></uris-resolved>
                    <availability><xsl:value-of select="$examples-are-available"/></availability>
                    <xsl:sequence use-when="$advanced-functions-available">
                        <raw-archives><xsl:copy-of select="$example-archives"/></raw-archives>
                        <archives-entries-maps><xsl:copy-of select="tan:map-to-xml($example-archive-entries-maps)"/></archives-entries-maps>
                        <first-archive-as-map-with-content><xsl:copy-of select="tan:map-to-xml(tan:extract-map($example-archives[1], $example-archive-entries-maps[1]))"/></first-archive-as-map-with-content>
                        <archive-keys><xsl:copy-of select="$example-archive-map-keys"/></archive-keys>
                    </xsl:sequence>
                    <components><xsl:copy-of select="$example-components"/></components>
                    <components-redirected><xsl:copy-of select="$redirected-components"/></components-redirected>
                    <components-revised><xsl:copy-of select="$revised-components"/></components-revised>
                </diagnostics>
            </xsl:when>
            <xsl:when test="true()" use-when="$advanced-functions-available">
                <xsl:for-each-group select="$revised-components" group-by="/*/@xml:base">
                    <xsl:variable name="new-target-uri" select="current-grouping-key()"/>
                    <xsl:sequence select="tan:save-archive(current-group(), $new-target-uri)"/>
                </xsl:for-each-group>
            </xsl:when>
            <xsl:otherwise>
                <xsl:document>
                    <xsl:for-each-group select="$revised-components" group-by="/*/@xml:base">
                        <xsl:variable name="new-target-uri" select="current-grouping-key()"/>
                        <xsl:call-template name="tan:save-archive">
                            <xsl:with-param name="archive-components" select="current-group()"/>
                            <xsl:with-param name="resolved-uri" select="$new-target-uri"/>
                        </xsl:call-template>
                    </xsl:for-each-group>
                </xsl:document>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
</xsl:stylesheet>
