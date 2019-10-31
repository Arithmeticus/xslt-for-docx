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
    
    <xsl:variable name="master-static-base-uri" select="static-base-uri()"/>

    <xsl:import href="../open-and-save-docx.xsl"/>
    <!-- The following inclusion has components common to stylesheets used for other examples. -->
    <xsl:include href="../inclusions/examples%20core.xsl"/>
    
    <!-- This stylesheet illustrates how to to change a Word or Excel file with regular expressions. -->
    <!-- Catalyzing input: any XML file (including this very one) -->
    <!-- Main input: the documents referred to by $example-urls -->
    <!-- Primary output: none, unless diagnostics are on -->
    <!-- Output: each document, saved under the output subdirectory -->

    <!-- Regular expressions are quite deficient in Microsoft Word's equivalent, wildcards, and 
        non-existent in Excel. This stylesheet opens up both formats to sophisticated changes 
        based on regular expressions. The key limitation of this application is that searches are made 
        exclusively within a single paragraph, and the replacement features apply exclusively to the 
        shared strings within an Excel file, not to values or formulas. (To extend the application
        to avoid those shortcomings would be straightforward, but I opted not to, in the interests of
        time.) -->

    <!-- This stylesheet has been developed in concert with the accompanying docx and xlsx 
        files. Word and Excel files can get quite complicated, and the work you wish to do may require 
        extra programming beyond the initial ideas outlined here. -->

    <xsl:output indent="yes" use-when="$diagnostics-on"/>
    
    <xsl:variable name="example-urls" as="element()*">
        <docx href="example-a.docx"/>
        <xlsx href="example-b.xlsx"/>
    </xsl:variable>
    
    <!-- The following three parameters are the ones traditionally associated with fn:replace(), and will behave the
    same way on the docx/xlsx file. -->
    <xsl:param name="pattern" as="xs:string">[bd-g]</xsl:param>
    <xsl:param name="replacement" as="xs:string" select="'!'"/>
    <xsl:param name="flags" as="xs:string?" select="'i'"/>


    <!-- PASS 1: IMPLANT PLAIN TEXT IN EACH PARAGRAPH NODE, CHANGE @XML:BASE -->
    
    <xsl:variable name="new-document-components-pass-1" as="document-node()*">
        <xsl:apply-templates select="$example-components" mode="alter-components"/>
    </xsl:variable>
    
    <xsl:template match="/*" mode="alter-components">
        <xsl:copy>
            <xsl:copy-of select="@* except @xml:base"/>
            <xsl:attribute name="xml:base" select="replace(@xml:base, '(/[^/]+)$', '/output$1')"
            />
            <xsl:apply-templates mode="#current"/>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="ssh:sst/ssh:si/ssh:t/text()" mode="alter-components">
        <xsl:choose>
            <xsl:when test="string-length($flags) gt 0">
                <xsl:value-of select="replace(., $pattern, $replacement, $flags)"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="replace(., $pattern, $replacement)"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    <xsl:template match="w:p" mode="alter-components">
        <!-- Splintered Seas Technique. See discussion in ../inclusions/examples%20core.xsl just 
            before tan:enumerate-characters() -->
        <xsl:variable name="this-plain-text">
            <xsl:apply-templates select="w:r" mode="archive-to-plain-text"/>
        </xsl:variable>
        <xsl:variable name="text-replaced" as="element()">
            <plain-text>
                <xsl:choose>
                    <xsl:when test="string-length($flags) gt 0">
                        <xsl:analyze-string select="$this-plain-text" regex="{$pattern}" flags="{$flags}">
                            <xsl:matching-substring>
                                <replace with="{replace(., $pattern, $replacement, $flags)}">
                                    <xsl:value-of select="."/>
                                </replace>
                            </xsl:matching-substring>
                            <xsl:non-matching-substring>
                                <keep>
                                    <xsl:value-of select="."/>
                                </keep>
                            </xsl:non-matching-substring>
                        </xsl:analyze-string>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:analyze-string select="$this-plain-text" regex="{$pattern}">
                            <xsl:matching-substring>
                                <replace with="{replace(., $pattern, $replacement)}">
                                    <xsl:value-of select="."/>
                                </replace>
                            </xsl:matching-substring>
                            <xsl:non-matching-substring>
                                <keep>
                                    <xsl:value-of select="."/>
                                </keep>
                            </xsl:non-matching-substring>
                        </xsl:analyze-string>
                    </xsl:otherwise>
                </xsl:choose>
            </plain-text>
        </xsl:variable>
        <xsl:variable name="text-replacement-enumerated"
            select="tan:enumerate-characters($text-replaced)"/>
        <xsl:variable name="text-replacement-with-insertions" as="element()*">
            <xsl:apply-templates select="$text-replacement-enumerated" mode="add-insertion"/>
        </xsl:variable>
        
        <xsl:variable name="self-with-characters-enumerated"
            select="tan:enumerate-characters(.)"/>
        
        <xsl:if test="$diagnostics-on">
            <xsl:comment select="'Diagnostics on'"/>
            <plain-text><xsl:value-of select="$this-plain-text"/></plain-text>
            <text-replaced><xsl:copy-of select="$text-replaced"/></text-replaced>
            <replacement-pass-1><xsl:copy-of select="$text-replacement-enumerated"/></replacement-pass-1>
            <replacement-pass-2><xsl:copy-of select="$text-replacement-with-insertions"/></replacement-pass-2>
            <self-enumerated><xsl:copy-of select="$self-with-characters-enumerated"/></self-enumerated>
        </xsl:if>
        
        <xsl:copy>
            <xsl:copy-of select="@*"/>
            <xsl:apply-templates select="$self-with-characters-enumerated/*"
                mode="restore-splintered-sea">
                <xsl:with-param name="characters-to-delete" tunnel="yes"
                    select="$text-replacement-with-insertions/replace/c"/>
                <xsl:with-param name="items-to-insert" tunnel="yes"
                    select="$text-replacement-with-insertions/replace/insertion"
                />
            </xsl:apply-templates>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="replace[c[@n]]" mode="add-insertion">
        <xsl:variable name="this-insertion" as="item()*" select="xs:string(@with)"/>
        <xsl:copy>
            <xsl:copy-of select="@*"/>
            <xsl:if test="exists($this-insertion)">
                <insertion>
                    <!-- Copy the first @n, to specify where the insertion should be placed -->
                    <xsl:copy-of select="c[1]/@n"/>
                    <xsl:copy-of select="$this-insertion"/>
                </insertion>
            </xsl:if>
            <xsl:copy-of select="node()"/>
        </xsl:copy>
    </xsl:template>
    
    
    <!-- PASS 2: CLEAN UP FILES -->
    
    <xsl:variable name="new-document-components-pass-2" as="document-node()*">
        <xsl:apply-templates select="$new-document-components-pass-1" mode="clean-docx"/>
    </xsl:variable>
    

    <!-- PASS 3: CHANGE @XML:BASE FOR OUTPUT URL -->
    
    <xsl:variable name="new-document-components-pass-3" as="document-node()*">
        <xsl:apply-templates select="$new-document-components-pass-2" mode="reset-xml-base"/>
    </xsl:variable>
    
    

    <xsl:template match="/">
        <xsl:choose>
            <xsl:when test="$diagnostics-on">
                <diagnostics>
                    <input><xsl:copy-of select="$example-components"/></input>
                    <output-pass-1><xsl:copy-of select="$new-document-components-pass-1"/></output-pass-1>
                    <output-pass-2><xsl:copy-of select="$new-document-components-pass-2"/></output-pass-2>
                    <output-pass-3><xsl:copy-of select="$new-document-components-pass-3"/></output-pass-3>
                </diagnostics>
            </xsl:when>
            <xsl:when test="true()" use-when="$advanced-functions-available">
                <xsl:for-each-group select="$new-document-components-pass-3" group-by="/*/@xml:base">
                    <xsl:variable name="new-target-uri" select="current-grouping-key()"/>
                    <xsl:sequence select="tan:save-archive(current-group(), $new-target-uri)"/>
                </xsl:for-each-group>
            </xsl:when>
            <xsl:otherwise>
                <xsl:for-each-group select="$new-document-components-pass-3" group-by="/*/@xml:base">
                    <xsl:variable name="new-target-uri" select="current-grouping-key()"/>
                    <xsl:call-template name="tan:save-archive">
                        <xsl:with-param name="archive-components" select="current-group()"/>
                        <xsl:with-param name="resolved-uri" select="$new-target-uri"/>
                    </xsl:call-template>
                </xsl:for-each-group>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
</xsl:stylesheet>
