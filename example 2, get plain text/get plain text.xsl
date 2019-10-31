<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:arch="http://expath.org/ns/archive" xmlns:file="http://expath.org/ns/file"
    xmlns:gpo="http://www.gpo.gov" xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:tan="tag:textalign.net,2015:ns" 
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:prop="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    xmlns:ssh="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    exclude-result-prefixes="#all" version="3.0">

    <xsl:param name="diagnostics-on" select="false()" static="true"/>

    <xsl:import href="../open-and-save-archive.xsl"/>
    <!-- The following inclusion has components common to stylesheets used for other examples. -->
    <xsl:include href="../inclusions/examples%20core.xsl"/>
    
    <!-- This stylesheet illustrates how to extract plain text from Word and Excel files. -->
    <!-- Catalyzing input: any XML file (including this very one) -->
    <!-- Main input: the documents referred to by $example-urls -->
    <!-- Primary output: the plain text for the docx/xlsx, concatenated -->
    <!-- Secondary output: none (but the stylesheet could easily be modified to generate secondary output) -->
    
    <!-- Output should be identical for both Saxon HE and Saxon PE+ scenarios -->
    
    <!-- This stylesheet has been developed solely for the accompanying examples. Docx and xlsx files can get 
        quite complicated, and your test case may require extra programming beyond the initial ideas outlined here. -->

    <xsl:output indent="yes" use-when="$diagnostics-on" use-character-maps="invisible-character-expansion"/>
    <xsl:output method="text" use-when="not($diagnostics-on)"/>
    
    <xsl:character-map name="invisible-character-expansion">
        <xsl:output-character character="&#x9;" string="&#x9;[TAB]"/>
        <xsl:output-character character="&#xad;" string="&#xad;[SHY]"/>
        <xsl:output-character character="&#x2011;" string="&#x2011;[NBH]"/>
    </xsl:character-map>

    <xsl:variable name="master-static-base-uri" select="static-base-uri()"/>

    <xsl:variable name="example-urls" as="element()*">
        <docx href="example-a.docx"/>
        <xlsx href="example-a.xlsx"/>
    </xsl:variable>
    
    <xsl:param name="keep-docx-endnotes" as="xs:boolean" select="true()"/>
    <xsl:param name="keep-docx-footnotes" as="xs:boolean" select="true()"/>
    <xsl:param name="keep-docx-headers" as="xs:boolean" select="true()"/>
    <xsl:param name="keep-docx-footers" as="xs:boolean" select="true()"/>
     
    <!-- items that you might want to suppress...or not -->
    <xsl:template match="w:endnotes" mode="archive-to-plain-text">
        <xsl:if test="$keep-docx-endnotes">
            <xsl:apply-templates mode="#current"/>
        </xsl:if>
    </xsl:template>
    <xsl:template match="w:footnotes" mode="archive-to-plain-text">
        <xsl:if test="$keep-docx-footnotes">
            <xsl:apply-templates mode="#current"/>
        </xsl:if>
    </xsl:template>
    <xsl:template match="w:hdr" mode="archive-to-plain-text">
        <xsl:if test="$keep-docx-headers">
            <xsl:apply-templates mode="#current"/>
        </xsl:if>
    </xsl:template>
    <xsl:template match="w:ftr" mode="archive-to-plain-text">
        <xsl:if test="$keep-docx-footers">
            <xsl:apply-templates mode="#current"/>
        </xsl:if>
    </xsl:template>

    <!-- INITIAL TEMPLATE -->
    <xsl:template match="/">
        <xsl:choose>
            <xsl:when test="$diagnostics-on">
                <diagnostics>
                    <input-components><xsl:copy-of select="$example-components"/></input-components>
                    <output><xsl:apply-templates select="$example-components" mode="archive-to-plain-text"/></output>
                </diagnostics>
            </xsl:when>
            <xsl:otherwise>
                <xsl:apply-templates select="$example-components" mode="archive-to-plain-text"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
</xsl:stylesheet>
