<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:tan="tag:textalign.net,2015:ns"
    exclude-result-prefixes="#all"
    version="2.0">
    <xsl:include href="../open-and-save-docx.xsl"/>
    
    <!-- This spreadsheet illustrates how to use xslt-for-docx, by showing how a Word file can be retrieved through XSLT, transformed, and then reassembled into a different version. -->
    <!-- Input: any XML file that has an element with an @href that points to a docx file -->
    <!-- Output: a copy of the docx file, with the contents changed by a regular expression defined in the parameters below; a time-date stamp is added to the filename of the result document -->
    <!-- This spreadsheet has been tested successfully on Saxon HE 9.6.0.7 via oXygen XML editor 18.1 -->
    
    <xsl:param name="pattern" as="xs:string?" select="'[aeiou]'"/>
    <xsl:param name="replacement" as="xs:string" select="''"/>
    <xsl:param name="flags" as="xs:string?" select="()"/>
    
    
    <xsl:variable name="first-location-with-docx"
        select="(//*[matches(@href, '\.docx$')][tan:docx-file-available(.)])[1]"/>
    <xsl:variable name="first-docx" select="tan:open-docx($first-location-with-docx)"/>
    <xsl:variable name="first-docx-changed" as="document-node()*">
        <xsl:for-each select="$first-docx">
            <xsl:choose>
                <xsl:when test="matches(*/@jar-path, '^word')">
                    <!-- if it's a main component of the Word document, e.g., word/document.xml, then perform the action -->
                    <xsl:document>
                        <xsl:apply-templates mode="replace-text"/>
                    </xsl:document>
                </xsl:when>
                <xsl:otherwise>
                    <!-- otherwise, keep the component intact -->
                    <xsl:copy-of select="."/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:for-each>
    </xsl:variable>
    
    <xsl:variable name="new-docx-uri"
        select="resolve-uri(concat(replace($first-location-with-docx/@href, '\.docx$', ''), replace(string(current-dateTime()), '\D', ''), '.docx'))"
    />

    <xsl:template match="/*">
        <xsl:call-template name="tan:save-docx">
            <xsl:with-param name="docx-parts" select="$first-docx-changed"/>
            <xsl:with-param name="resolved-uri" select="$new-docx-uri"/>
        </xsl:call-template>
    </xsl:template>

    <xsl:template match="comment()|processing-instruction()" mode="replace-text">
        <xsl:copy-of select="."/>
    </xsl:template>
    <xsl:template match="*" mode="replace-text">
        <xsl:copy>
            <xsl:copy-of select="@*"/>
            <xsl:apply-templates mode="#current"/>
        </xsl:copy>
    </xsl:template>
    <xsl:template match="text()" mode="replace-text">
        <xsl:choose>
            <xsl:when test="string-length($flags) gt 0">
                <xsl:value-of select="replace(., $pattern, $replacement, $flags)"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="replace(., $pattern, $replacement)"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
</xsl:stylesheet>