<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:arch="http://expath.org/ns/archive" xmlns:file="http://expath.org/ns/file"
    xmlns:gpo="http://www.gpo.gov" xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:tan="tag:textalign.net,2015:ns" exclude-result-prefixes="#all" version="3.0">

    <!-- This stylesheet illustrates how to extract the components of an archive, then repackage them and save the archive somewhere else. -->
    <!-- Catalyzing input: any XML file (including this very one) -->
    <!-- Main input: the documents referred to by $example-urls -->
    <!-- Primary output: none, unless diagnostics are on -->
    <!-- Secondary output: each document, saved in the output folder, and the components for each file expanded to subdirectories of output -->
    
    <xsl:param name="diagnostics-on" select="false()" static="true"/>
    
    <xsl:import href="../open-and-save-archive.xsl"/>
    <!-- The following inclusion has components common to stylesheets used for other examples. -->
    <xsl:include href="../inclusions/examples%20core.xsl"/>

    <xsl:output indent="yes" use-when="$diagnostics-on"/>
    <xsl:output indent="no" omit-xml-declaration="yes" use-when="not($diagnostics-on)"/>

    <xsl:variable name="master-static-base-uri" select="static-base-uri()"/>

    <xsl:variable name="example-urls" as="element()*">
        <zip href="example-a.zip"/>
        <zip href="example-b.zip"/>
        <zip href="example-a+b.zip"/>
        <zip href="example-a+b+c.zip"/>
        <docx href="example-d.docx"/>
        <docx href="example-e.docx"/>
        <docx href="example-f.docx"/>
        <xlsx href="example-g.xlsx"/>
        <xlsx href="example-h.xlsx"/>
        <odt href="example-i.odt"/>
        <ods href="example-j.ods"/>
        <sxc href="example-k.sxc"/>
        <sxw href="example-l.sxw"/>
        <epub href="example-m.epub"/>
        <jar href="example-n.jar"/>
    </xsl:variable>

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
                    <components-prepped-for-saving><xsl:apply-templates select="$example-components" mode="save-components-locally"/></components-prepped-for-saving>
                </diagnostics>
            </xsl:when>
            <xsl:when test="true()" use-when="$advanced-functions-available">
                <xsl:apply-templates select="$example-components" mode="save-components-locally"/>
                <xsl:for-each-group select="$redirected-components" group-by="/*/@xml:base">
                    <xsl:variable name="new-target-uri" select="current-grouping-key()"/>
                    <xsl:sequence select="tan:save-archive(current-group(), $new-target-uri)"/>
                </xsl:for-each-group>
            </xsl:when>
            <xsl:otherwise>
                <xsl:document>
                    <xsl:apply-templates select="$example-components" mode="save-components-locally"/>
                    <xsl:for-each-group select="$redirected-components" group-by="/*/@xml:base">
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
