<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:arch="http://expath.org/ns/archive" xmlns:file="http://expath.org/ns/file"
    xmlns:gpo="http://www.gpo.gov" xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:tan="tag:textalign.net,2015:ns"
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:prop="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    xmlns:ssh="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    exclude-result-prefixes="#all" version="3.0">

    <xsl:param name="diagnostics-on" select="false()" static="true"/>

    <xsl:import href="../open-and-save-docx.xsl"/>
    <!-- The following inclusion has components common to stylesheets used for other examples. -->
    <xsl:include href="../inclusions/examples%20core.xsl"/>
    
    <!-- This stylesheet illustrates how to generate form letters based upon an XML database and a template Word file. -->
    <!-- Catalyzing input: any XML file (including this very one) -->
    <!-- Main input: the documents bound to $donor-letter-components and $donor-database -->
    <!-- Primary output: none, unless diagnostics are on -->
    <!-- Secondary output: one docx file per donor, populated with the appropriate data, saved in the output directory -->

    <!-- This application replicates the traditional mailings method built into Microsoft Word, and even improves on the model:
        - You can put your data in an XML file or a format you like.
        - One-to-many relationships between a record and its data is supported (i.e., you don't need to try to fit it into a one-to-one data structure like an Excel spreadsheet).
        - You can add whatever variables you like by defining them here, then typing them in the Word document where you want, styled as you like (in this example $ signals a variable).
        - You can customize and transform the data returned to a variable with rather complex XPath functions. See tan:resolve-variable() -->

    <!-- This stylesheet has been developed solely for the accompanying examples. Word files can get 
        quite complicated, and your test case may require extra programming beyond the initial ideas outlined here. -->

    <xsl:output indent="yes" use-when="$diagnostics-on"/>
    
    <xsl:variable name="master-static-base-uri" select="static-base-uri()"/>

    <!-- How can we tell, in a Word document, whether we've found a variable? -->
    <xsl:variable name="variable-name-regex" as="xs:string">\$([\p{L}\d_-]+)</xsl:variable>
    <!-- I've thrown this variable in, to illustrate how the source data could be filtered, to avoid certain records -->
    <xsl:variable name="year-documented" select="'2003'"/>

    <xsl:variable name="example-urls" as="element()?">
        <docx href="donor%20letter.docx"/>
    </xsl:variable>

    <xsl:variable name="donor-letter-uri-resolved"
        select="resolve-uri($example-urls[1]/@href, static-base-uri())"/>
    <xsl:variable name="donor-letter-components" select="tan:open-docx($donor-letter-uri-resolved)"/>
    <xsl:variable name="donor-database-uri" select="resolve-uri('donors.xml', static-base-uri())"/>
    <xsl:variable name="donor-database" select="doc($donor-database-uri)"/>
    
    <xsl:function name="tan:resolve-variable" as="item()*">
        <!-- Input: the name of a variable (lowercase); nodes that provide context -->
        <!-- Output: items corresponding to the variable name, as found on the context nodes -->
        <!-- If a variable is not defined, nothing is returned -->
        <!-- Variables are considered not case-sensitive, so name-matching is done on the basis of lowercase -->
        <xsl:param name="variable-name" as="xs:string"/>
        <xsl:param name="context-nodes" as="node()*"/>
        <xsl:choose>
            <xsl:when test="$variable-name = 'donorfullname'">
                <!-- Note, if a record in the data has more than one <name>, it'll all get squashed together. -->
                <xsl:value-of select="$context-nodes//name"/>
            </xsl:when>
            <xsl:when test="$variable-name = 'donoraddress'">
                <!-- Here we know that there will be more than one line. We could have used an iterate command
                    (see below) but I wanted to show that you could do simple iteration within the XSLT document, 
                    assuming we don't need new paragraphs. So in this case, we iterate over lines in the address, 
                    but we insert soft carriage returns. -->
                <xsl:variable name="address-lines" select="$context-nodes//address/line"/>
                <xsl:for-each select="$address-lines">
                    <xsl:value-of select="normalize-space(.)"/>
                    <xsl:if test="position() lt count($address-lines)">
                        <w:br/>
                    </xsl:if>
                </xsl:for-each>
            </xsl:when>
            <xsl:when test="$variable-name = 'donortitle'">
                <xsl:value-of select="$context-nodes//title"/>
            </xsl:when>
            <xsl:when test="$variable-name = 'donorfamilyname'">
                <xsl:value-of select="$context-nodes//name/family"/>
            </xsl:when>
            <xsl:when test="$variable-name = 'donorgivenname'">
                <xsl:value-of select="$context-nodes//name/given"/>
            </xsl:when>
            <xsl:when test="$variable-name = 'emailaddresses'">
                <!-- This case demonstrates how the data can be finessed. In this case, we want 
                to make sure that multiple values are correctly punctuated and phrased. -->
                <xsl:value-of select="tan:commas-and-ands($context-nodes//email)"/>
            </xsl:when>
            <xsl:when test="$variable-name = 'emailqty'">
                <xsl:value-of select="count($context-nodes//email)"/>
            </xsl:when>
            <xsl:when test="$variable-name = 'plemailqty'">
                <!-- Here we can add an 'es' to 'address' if there are multiple email addresses. -->
                <xsl:choose>
                    <xsl:when test="count($context-nodes//email) gt 1">
                        <xsl:value-of select="'es'"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:value-of select="''"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:when test="$variable-name = 'giftdate'">
                <!-- So the raw data is in ISO format (as it should be). This shows how we can
                change the date format on the fly. So much easier to do with XPath functions 
                instead of Word! -->
                <xsl:value-of
                    select="
                        for $i in $context-nodes//date
                        return
                            format-date($i, '[MNn] [D], [Y]')"
                />
            </xsl:when>
            <xsl:when test="$variable-name = 'giftamount'">
                <!-- Again, we can change not just the format of dates but of numbers. -->
                <xsl:value-of
                    select="
                        for $i in $context-nodes//amount
                        return
                            format-number($i, '$0.00')"
                />
            </xsl:when>
            <xsl:when test="$variable-name = 'giftcomment'">
                <xsl:value-of select="$context-nodes//comment"/>
            </xsl:when>
            <xsl:when test="$variable-name = 'gifttotal'">
                <!-- Here's some simple addition, done on the fly, but we could do all sorts
                of more complex mathematics, if we wanted to. -->
                <xsl:variable name="these-amounts"
                    select="
                        for $i in $context-nodes//amount
                        return
                            number($i)"
                />
                <xsl:value-of select="format-number(sum($these-amounts), '$0.00')"/>
            </xsl:when>
            <xsl:otherwise>
                <!-- Let the user know if a variable isn't understood -->
                <xsl:message select="'variable ' || $variable-name || ' undefined'"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:function>
    
    <xsl:function name="tan:commas-and-ands" as="xs:string?">
        <!-- One-parameter version of the full one below -->
        <xsl:param name="input-strings" as="xs:string*"/>
        <xsl:value-of select="tan:commas-and-ands($input-strings, true())"/>
    </xsl:function>
    <xsl:function name="tan:commas-and-ands" as="xs:string?">
        <!-- Input: sequences of strings; a boolean -->
        <!-- Output: the strings joined together with , and 'and'. If the 2nd parameter is true, then
        the Oxford (serial) comma is used (e.g., A, B, and C instead of A, B and C)-->
        <xsl:param name="input-strings" as="xs:string*"/>
        <xsl:param name="oxford-comma" as="xs:boolean"/>
        <xsl:variable name="input-string-count" select="count($input-strings)"/>
        <xsl:variable name="results" as="xs:string*">
            <xsl:for-each select="$input-strings">
                <xsl:variable name="this-pos" select="position()"/>
                <xsl:value-of select="."/>
                <xsl:if test="$input-string-count gt 2">
                    <xsl:choose>
                        <xsl:when test="$this-pos lt ($input-string-count - 1)">,</xsl:when>
                        <xsl:when test="$this-pos = ($input-string-count - 1) and $oxford-comma">,</xsl:when>
                    </xsl:choose>
                </xsl:if>
                <xsl:if test="$this-pos lt $input-string-count">
                    <xsl:text> </xsl:text>
                </xsl:if>
                <xsl:if test="$input-string-count gt 1 and $this-pos = ($input-string-count - 1)"
                    >and </xsl:if>
            </xsl:for-each>
        </xsl:variable>
        <xsl:value-of select="string-join($results)"/>
    </xsl:function>


    <!-- PASS 1: TEMPORARILY IMPLANT PLAIN TEXT IN EACH PARAGRAPH NODE, CHANGE @XML:BASE, TEMPORARILY IMPLANT DONOR DATA -->
    
    <xsl:variable name="new-document-components-pass-1" as="document-node()*">
        <xsl:for-each select="$donor-database/donors/donor">
            <xsl:variable name="this-donor-data" select="."/>
            <xsl:apply-templates select="$donor-letter-components"
                mode="prep-letter">
                <xsl:with-param name="donor-data" select="$this-donor-data" tunnel="yes"/>
            </xsl:apply-templates>
        </xsl:for-each>
    </xsl:variable>
    
    <xsl:template match="/*" mode="prep-letter">
        <xsl:param name="donor-data" as="element()+" tunnel="yes"/>
        <xsl:copy>
            <xsl:copy-of select="@* except @xml:base"/>
            <xsl:attribute name="xml:base"
                select="replace(@xml:base, 'donor%20letter', 'output/donor%20letter-' || lower-case($donor-data/id) || $output-url-infix)"/>
            <xsl:if test="self::w:document">
                <xsl:apply-templates select="$donor-data" mode="#current"/>
            </xsl:if>
            <xsl:apply-templates mode="#current"/>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="donation[date]" mode="prep-letter">
        <xsl:variable name="matching-date" select="date[matches(., '^' || $year-documented)]"/>
        <xsl:if test="exists($matching-date)">
            <xsl:copy-of select="."/>
        </xsl:if>
    </xsl:template>
    
    <xsl:template match="w:p" mode="prep-letter">
        <xsl:copy>
            <xsl:copy-of select="@*"/>
            <plain-text>
                <xsl:apply-templates select="w:r" mode="archive-to-plain-text"/>
            </plain-text>
            <xsl:apply-templates mode="#current"/>
        </xsl:copy>
    </xsl:template>
    
    
    <!-- PASS 2: POPULATE THE LETTER -->
    
    <xsl:variable name="new-document-components-pass-2" as="document-node()*">
        <xsl:apply-templates select="$new-document-components-pass-1" mode="populate-letter"/>
    </xsl:variable>
    
    <xsl:template match="w:document" mode="populate-letter">
        <xsl:copy>
            <xsl:copy-of select="@*"/>
            <xsl:apply-templates mode="#current">
                <xsl:with-param name="donor-data" tunnel="yes" select="donor"/>
            </xsl:apply-templates>
        </xsl:copy>
    </xsl:template>
    
    <!-- The next variable defines the pattern for the iteration command. In parentheses is expected the name
    of a descendant node in the XML data for the record at hand. -->
    <xsl:variable name="next-para-iterate-regex" as="xs:string">!nextParaForEach\(([\p{L}_-]+)\)</xsl:variable>
    <xsl:template match="w:p" mode="populate-letter">
        <!-- Employs Splintered Seas Technique. See discussion in ../inclusions/examples%20core.xsl just 
            before tan:enumerate-characters() -->
        <!-- This is an important template. If the <p> has a command to iterate over the next paragraph, skip it. If 
            the <p> is preceded by an iteration <p>, repeat it for every node specified in the iteration command. 
            Whether treated as a single copy or an iteration, the <p>'s <plain-text> is scanned for variables.
            If there are no variables, the paragraph is left as is. If there are variables, then the Splintered Sea
            Technique (see above) is used on <plain-text> and the <p>, with <plain-text> being assessed
            ahead of time for its variable content.
        -->
        <xsl:variable name="this-p" select="."/>
        <xsl:variable name="omit-this-p"
            select="matches(plain-text, $next-para-iterate-regex)"/>
        <xsl:choose>
            <xsl:when test="$omit-this-p"/>
            <xsl:otherwise>
                <xsl:variable name="prev-p" select="preceding-sibling::w:p[1]"/>
                <xsl:variable name="prev-p-iteration-command" as="element()?">
                    <xsl:analyze-string select="$prev-p/plain-text"
                        regex="{$next-para-iterate-regex}">
                        <xsl:matching-substring>
                            <iterate on="{regex-group(1)}"/>
                        </xsl:matching-substring>
                    </xsl:analyze-string>
                </xsl:variable>
                <xsl:variable name="nodes-to-iterate-on"
                    select="root(.)/w:document/donor//*[name(.) = $prev-p-iteration-command/@on]"
                />
                <xsl:variable name="self-with-characters-enumerated"
                    select="tan:enumerate-characters(.)"/>
                <xsl:variable name="text-checked-for-variables" as="element()">
                    <variable-check>
                        <xsl:analyze-string select="plain-text"
                            regex="{$variable-name-regex}">
                            <xsl:matching-substring>
                                <!-- variable names are not case-sensitive -->
                                <variable name="{lower-case(regex-group(1))}">
                                    <xsl:value-of select="."/>
                                </variable>
                            </xsl:matching-substring>
                            <xsl:non-matching-substring>
                                <text>
                                    <xsl:value-of select="."/>
                                </text>
                            </xsl:non-matching-substring>
                        </xsl:analyze-string>
                    </variable-check>
                </xsl:variable>
                <xsl:variable name="text-checked-for-variables-with-characters-enumerated"
                    select="tan:enumerate-characters($text-checked-for-variables)"/>

                <xsl:if test="$diagnostics-on">
                    <plain-text-splintered-sea><xsl:copy-of select="$text-checked-for-variables-with-characters-enumerated"/></plain-text-splintered-sea>
                    <this-p-splintered-sea><xsl:copy-of select="$self-with-characters-enumerated"/></this-p-splintered-sea>
                </xsl:if>
                <xsl:choose>
                    
                    <!-- If we need to repeat the current paragraph a certain number of times, we treat that specially -->
                    <xsl:when test="exists($nodes-to-iterate-on)">
                        <xsl:for-each select="$nodes-to-iterate-on">
                            <xsl:variable name="this-data-node" select="."/>
                            <xsl:variable
                                name="text-checked-for-variables-with-characters-enumerated-and-substitutions-marked"
                                as="element()*">
                                <xsl:apply-templates
                                    select="$text-checked-for-variables-with-characters-enumerated"
                                    mode="add-insertion">
                                    <xsl:with-param name="donor-data" tunnel="yes"
                                        select="$this-data-node"/>
                                </xsl:apply-templates>
                            </xsl:variable>
                            <xsl:if test="$diagnostics-on">
                                <plain-text-splintered-sea-prepped><xsl:copy-of select="$text-checked-for-variables-with-characters-enumerated-and-substitutions-marked"/></plain-text-splintered-sea-prepped>
                            </xsl:if>
                            <w:p>
                                <xsl:copy-of select="$this-p/@*"/>
                                <xsl:apply-templates select="$self-with-characters-enumerated/*"
                                    mode="restore-splintered-sea">
                                    <xsl:with-param name="characters-to-delete" tunnel="yes"
                                        select="$text-checked-for-variables-with-characters-enumerated-and-substitutions-marked/variable[insertion]/c"/>
                                    <xsl:with-param name="items-to-insert" tunnel="yes"
                                        select="$text-checked-for-variables-with-characters-enumerated-and-substitutions-marked/variable/insertion"
                                    />
                                </xsl:apply-templates>
                            </w:p>
                        </xsl:for-each>
                    </xsl:when>
                    
                    <!-- Otherwise we make straight-forward substitutions. -->
                    <xsl:otherwise>
                        <xsl:variable
                            name="text-checked-for-variables-with-characters-enumerated-and-substitutions-marked"
                            as="element()*">
                            <xsl:apply-templates
                                select="$text-checked-for-variables-with-characters-enumerated"
                                mode="add-insertion"/>
                        </xsl:variable>
                        <xsl:if test="$diagnostics-on">
                            <plain-text-splintered-sea-prepped><xsl:copy-of select="$text-checked-for-variables-with-characters-enumerated-and-substitutions-marked"/></plain-text-splintered-sea-prepped>
                        </xsl:if>
                        <xsl:copy>
                            <xsl:copy-of select="@*"/>
                            <xsl:apply-templates select="$self-with-characters-enumerated/*"
                                mode="restore-splintered-sea">
                                <xsl:with-param name="characters-to-delete" tunnel="yes"
                                    select="$text-checked-for-variables-with-characters-enumerated-and-substitutions-marked/variable[insertion]/c"/>
                                <xsl:with-param name="items-to-insert" tunnel="yes"
                                    select="$text-checked-for-variables-with-characters-enumerated-and-substitutions-marked/variable/insertion"
                                />
                            </xsl:apply-templates>
                        </xsl:copy>

                    </xsl:otherwise>
                </xsl:choose>

            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- These two templates make certain that //w:p/plain-text doesn't foul up the character count -->
    <xsl:template match="plain-text" mode="archive-to-plain-text"/>
    <xsl:template match="plain-text" priority="1" mode="enumerate-characters">
        <xsl:copy-of select="."/>
    </xsl:template>
    
    <xsl:template match="variable[c[@n]]" mode="add-insertion">
        <xsl:param name="donor-data" tunnel="yes" as="element()+"/>
        <xsl:variable name="this-variable-name" select="@name"/>
        <xsl:variable name="this-insertion" as="item()*" select="tan:resolve-variable($this-variable-name, $donor-data)"/>
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
    
    
    <!-- PASS 3: CLEAN UP FILES -->
    
    <xsl:variable name="new-document-components-pass-3" as="document-node()*">
        <xsl:apply-templates select="$new-document-components-pass-2" mode="clean-docx"/>
    </xsl:variable>
    

    <xsl:template match="/">
        <xsl:choose>
            <xsl:when test="$diagnostics-on">
                <diagnostics>
                    <xsl:copy-of select="$donor-database"/>
                    <letter-template-components><xsl:copy-of select="$donor-letter-components"/></letter-template-components>
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
