<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:arch="http://expath.org/ns/archive" 
    xmlns:map="http://www.w3.org/2005/xpath-functions/map"
    xmlns:file="http://expath.org/ns/file" xmlns:ssh="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:tan="tag:textalign.net,2015:ns" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:prop="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    exclude-result-prefixes="#all" version="3.0">

    <!-- stylesheet components common to the examples -->

    
    <!-- DEFAULTE TEMPLATE BEHAVIOR -->
    
    <!-- default template behavior is shallow copy, both nodes and attributes -->
    <xsl:template match="document-node()" priority="-2" mode="#all">
        <xsl:document>
            <xsl:apply-templates mode="#current"/>
        </xsl:document>
    </xsl:template>
    <xsl:template match="node() | @*" priority="-2" mode="#all">
        <xsl:copy>
            <xsl:apply-templates select="node() | @*" mode="#current"/>
        </xsl:copy>
    </xsl:template>

    <!-- IMPORTANT: The next template allows the imported open-and-save-docx.xsl to do what it's 
        supposed to do, without interference from the importing XSLT. -->
    <xsl:template match="document-node() | node() | @*" priority="2"
        mode="clean-up-archive map-to-xml">
        <xsl:apply-imports/>
    </xsl:template>
    
    
    
    <!-- GLOBAL VARIABLES -->
    
    <xsl:variable name="example-uris-resolved" as="xs:string*"
        select="
            for $i in $example-urls/@href
            return
                resolve-uri($i, $master-static-base-uri)"/>
    <xsl:variable name="examples-are-available" as="xs:boolean*"
        select="
            for $i in $example-uris-resolved
            return
                tan:docx-file-available($i)"/>

    <xsl:variable name="example-components"
        select="
            for $i in $example-uris-resolved
            return
                tan:open-archive($i)"/>
    
    <!-- unique to PE+ version -->
    <xsl:variable name="example-archives" as="xs:base64Binary*"
        use-when="$advanced-functions-available"
        select="
            for $i in $example-uris-resolved
            return
                tan:open-raw-archive($i)"
    />
    
    <!-- unique to PE+ version -->
    <xsl:variable name="example-entries" as="map(xs:string,map(xs:string,item()*))*"
        use-when="$advanced-functions-available"
        select="
            for $i in $example-archives
            return
                arch:entries-map($i)"
    />
    
    <!-- unique to PE+ version -->
    <xsl:variable name="example-archive-entries-maps" as="map(xs:string,map(xs:string,item()?))*"
        use-when="$advanced-functions-available"
        select="
            for $i in $example-uris-resolved
            return
                tan:entries-map($i)"
    />
    
    <!-- unique to PE+ version -->
    <xsl:variable name="example-archive-map-keys" use-when="$advanced-functions-available"
        select="
            for $i in $example-archive-entries-maps
            return
                map:keys($i)"
    />
    
    <xsl:template match="@xml:base" mode="reset-xml-base">
        <xsl:attribute name="{name(.)}">
            <xsl:value-of select="replace(replace(., '(\.[^.]+)$', $output-url-infix || '-output$1'),'/([^/]+[/!]*)$','/output/$1')"/>
        </xsl:attribute>
    </xsl:template>
    
    <xsl:variable name="redirected-components" as="document-node()*">
        <xsl:apply-templates select="$example-components" mode="reset-xml-base"/>
    </xsl:variable>
    
    <xsl:variable name="output-url-infix" as="xs:string"
        select="
            if ($advanced-functions-available) then
                '-saxon_pe+'
            else
                '-saxon_he'"
    />
    
    <xsl:variable name="characters-forbidden-in-fullpaths-regex"
        >[&lt;&gt;"\|\?\*\[\]\{\}]</xsl:variable>
    
    
    <!-- FUNCTIONS AND OTHER UTILITIES -->
    
    <!-- EXTRACTING EXCEL DATA -->
    <xsl:function name="tan:excel-column-number" as="xs:integer?">
        <!-- Input: a string representing an Excel column number; a boolean indicating whether the letter should be normalized -->
        <!-- Output: the integer corresponding to the letter -->
        <xsl:param name="column-letter" as="xs:string"/>
        <xsl:param name="normalize-letter" as="xs:boolean"/>
        <xsl:variable name="this-letter"
            select="
            if ($normalize-letter) then
            replace(upper-case($column-letter), '[^A-Z]', '')
            else
            $column-letter"
        />
        <xsl:variable name="is-valid-column-letter" select="string-length($this-letter) gt 0 and matches($this-letter, '^[A-X]?[A-Z]?[A-Z]$')"/>
        <xsl:choose>
            <xsl:when test="$is-valid-column-letter">
                <xsl:variable name="these-values-reversed"
                    select="
                    for $i in reverse(string-to-codepoints($this-letter))
                    return
                    $i - 64"
                />
                <xsl:variable name="new-value" as="xs:integer" select="($these-values-reversed[1]) + (($these-values-reversed[2], 0)[1] * 26) + (($these-values-reversed[3], 0)[1] * 26 * 26)"/>
                <xsl:value-of select="$new-value"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:message select="$column-letter || ' (' || $this-letter || ') is not a valid column letter'"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:function>
    
    <!-- CONVERSION TO PLAIN TEXT -->
    <xsl:template match="*" mode="archive-to-plain-text">
        <xsl:apply-templates mode="#current"/>
    </xsl:template>
    <xsl:template match="processing-instruction() | comment()"
        mode="archive-to-plain-text enumerate-characters create-splintered-sea"/>
    <xsl:template match="w:p" mode="archive-to-plain-text">
        <xsl:apply-templates mode="#current"/>
        <xsl:text>&#xa;</xsl:text>
    </xsl:template>
    <xsl:template match="ssh:c[not(@t)]" mode="archive-to-plain-text">
        <xsl:apply-templates mode="#current"/>
        <xsl:text>&#x9;</xsl:text>
    </xsl:template>
    <xsl:template match="w:tab" mode="archive-to-plain-text">
        <xsl:text>&#x9;</xsl:text>
    </xsl:template>
    <xsl:template match="w:br" mode="archive-to-plain-text">
        <xsl:text>&#xd;</xsl:text>
    </xsl:template>
    <xsl:template match="w:noBreakHyphen" mode="archive-to-plain-text">
        <xsl:text>&#x2011;</xsl:text>
    </xsl:template>
    <xsl:template match="w:softHyphen" mode="archive-to-plain-text">
        <xsl:text>&#xad;</xsl:text>
    </xsl:template>
    <xsl:template match="ssh:sheetData" mode="archive-to-plain-text">
        <xsl:variable name="this-base" select="root()/*/@xml:base"/>
        <xsl:variable name="this-sst" select="$example-components[ssh:sst/@xml:base = $this-base]"/>
        <xsl:apply-templates mode="#current">
            <xsl:with-param name="sst" tunnel="yes" as="document-node()?" select="$this-sst"/>
        </xsl:apply-templates>
    </xsl:template>
    <xsl:template match="ssh:row" mode="archive-to-plain-text">
        <xsl:variable name="this-row" select="."/>
        <xsl:variable name="these-span-numbers"
            select="
            for $i in tokenize(@spans, ':')
            return
            xs:integer($i)"
        />
        <xsl:variable name="expected-cell-qty" select="$these-span-numbers[2] - $these-span-numbers[1] + 1"/>
        <xsl:for-each select="1 to $expected-cell-qty">
            <!-- It's important to find the cell gaps within a given row and replace them with tabs. -->
            <xsl:variable name="this-int" select="."/>
            <xsl:variable name="this-c" select="$this-row/ssh:c[tan:excel-column-number(@r, true()) = $this-int]"/>
            <xsl:choose>
                <xsl:when test="exists($this-c)">
                    <xsl:apply-templates select="$this-c" mode="#current"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:text>&#x9;</xsl:text>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:for-each>
        <xsl:text>&#xa;</xsl:text>
    </xsl:template>
    <xsl:template match="ssh:c[@t = 's']" mode="archive-to-plain-text">
        <!-- These are cells that invoke the shared strings -->
        <xsl:param name="sst" tunnel="yes" as="document-node()"/>
        <xsl:variable name="this-sst-no" select="number(ssh:v)"/>
        <xsl:variable name="this-sst-val" select="$sst/ssh:sst/ssh:si[$this-sst-no + 1]/ssh:t"/>
        <xsl:value-of select="$this-sst-val"/>
        <xsl:if test="exists(following-sibling::ssh:c)">
            <xsl:text>&#x9;</xsl:text>
        </xsl:if>
    </xsl:template>
    <!-- items to suppress -->
    <xsl:template match="w:instrText | prop:Properties | cp:coreProperties | w:pPr" mode="archive-to-plain-text"/>
    <xsl:template match="ssh:f | ssh:sst | _base64Binary" mode="archive-to-plain-text"/>
    
    <!-- CHANGE CONTENT OF WORD FILES -->
    <!-- Splintered Seas Technique -->
    <!-- 
            How do you change text in a Word document? Very carefully. The text of every paragraph (<w:p>) is 
            in the text content of one or more ranges, <w:r>. Every time Word wants to mark a particular word or 
            letter, e.g., change of editor, change of font, a misspelled word, a new comment, the software splits 
            the appropriate <w:r>. You should assume that a particular phrase you're looking for is split across 
            <w:r>s.
                 My approach to changes the Word file employs an approach I'll call the Splintered Seas Technique, 
            because it relies upon taking the original <w:p> and putting each text codepoint into its own 
            <c n="[POSITION]"> - - splintering the text into <c>s - - then removing <c>s in a second pass, based upon
            whether an @n is to be retained, dropped, or replaced, normally by checking against a tunnel 
            parameter that itself is a splintered sea, or a derivative. -->
    
    <xsl:function name="tan:enumerate-characters" as="element()*">
        <!-- Input: a sequence of elements -->
        <!-- Output: those same elements, but each character in every text node wrapped with <c n="">, where @n
            is the master character position within the original input. -->
        <!-- This function is written primarily to begin the Splintered Seas Technique, which 
            allows one to find, and then strike or replace, a text distributed throughout a tree 
            fragment of unknown depth and length. On the SST, see comments at the example
            dealing with regular expressions. The Splintered Seas Technique concludes with the
            restore-splintered-sea template mode, below
        -->
        <xsl:param name="element-sequence" as="element()*"/>
        <xsl:apply-templates select="$element-sequence" mode="enumerate-characters">
            <xsl:with-param name="previous-character-number" select="0" tunnel="yes"/>
        </xsl:apply-templates>
    </xsl:function>
    
    <xsl:template match="*" mode="enumerate-characters create-splintered-sea">
        <xsl:param name="previous-character-number" tunnel="yes" as="xs:integer" select="0"/>
        <xsl:copy>
            <xsl:copy-of select="@*"/>
            <xsl:iterate select="node()">
                <xsl:param name="previous-char-no" as="xs:integer"
                    select="$previous-character-number"/>
                <xsl:variable name="this-plain-text">
                    <xsl:apply-templates select="." mode="archive-to-plain-text"/>
                </xsl:variable>
                <xsl:apply-templates select="." mode="#current">
                    <xsl:with-param name="previous-character-number" select="$previous-char-no" tunnel="yes"/>
                </xsl:apply-templates>
                <xsl:next-iteration>
                    <xsl:with-param name="previous-char-no"
                        select="$previous-char-no + string-length($this-plain-text)"/>
                </xsl:next-iteration>
            </xsl:iterate>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="text()" mode="enumerate-characters create-splintered-sea">
        <xsl:param name="previous-character-number" tunnel="yes" as="xs:integer" select="0"/>
        <xsl:for-each select="string-to-codepoints(.)">
            <c n="{position() + $previous-character-number}">
                <xsl:value-of select="codepoints-to-string(.)"/>
            </c>
        </xsl:for-each>
    </xsl:template>
    <xsl:template match="w:tab | w:noBreakHyphen | w:softHyphen | w:Br" priority="1" mode="enumerate-characters">
        <xsl:param name="previous-character-number" tunnel="yes" as="xs:integer" select="0"/>
        <c n="{$previous-character-number + 1}">
            <xsl:copy-of select="."/>
        </c>
    </xsl:template>
    <xsl:template match="w:pPr" priority="1" mode="enumerate-characters create-splintered-sea">
        <xsl:copy-of select="."/>
    </xsl:template>
    
    <!-- This template, the second half of the Splintered Seas Technique, reversing the 
        enumeration process above, deleting select <c>s, inserting select content, and shallow-skipping 
        all other <c>s.-->
    <xsl:template match="c" mode="restore-splintered-sea">
        <xsl:param name="characters-to-delete" tunnel="yes" as="element()*"/>
        <xsl:param name="items-to-insert" tunnel="yes" as="element()*"/>
        <xsl:variable name="this-n" select="@n"/>
        <xsl:variable name="these-items-to-insert" select="$items-to-insert[@n = $this-n]/node()"/>
        <xsl:variable name="delete-me" select="$this-n = $characters-to-delete/@n"/>
        <xsl:copy-of select="$these-items-to-insert"/>
        <xsl:if test="not($delete-me)">
            <xsl:apply-templates mode="#current"/>
        </xsl:if>
    </xsl:template>
    
    
    <!-- CLEAN UP WORD FILES -->
    <xsl:template match="w:p/w:r" mode="clean-docx">
        <xsl:variable name="has-good-ts" select="w:t[text()]"/>
        <xsl:variable name="has-alternate-content" select="exists(mc:*) or exists(w:drawing)"/>
        <xsl:variable name="has-special-chars" select="exists(w:tab) or exists(w:br) or exists(w:softHyphen) or exists(w:noBreakHyphen)"/>
        <xsl:if test="$has-good-ts or $has-alternate-content or $has-special-chars">
            <xsl:copy>
                <xsl:copy-of select="@*"/>
                <xsl:apply-templates mode="#current"/>
            </xsl:copy>
        </xsl:if>
    </xsl:template>
    <!-- delete empty <w:t>s and ad hoc additions -->
    <xsl:template match="w:t[not(text())] | plain-text | w:document/donor" mode="clean-docx"/>
    <xsl:template match="w:t[w:br]" mode="clean-docx">
        <xsl:apply-templates mode="#current"/>
    </xsl:template>
    <xsl:template match="w:t[w:br]/text()" mode="clean-docx">
        <w:t>
            <xsl:value-of select="."/>
        </w:t>
    </xsl:template>
    

    <!-- SAVE INDIVIDUAL COMPONENTS -->
    <xsl:template match="/" mode="save-components-locally">
        <xsl:variable name="this-target-base-normalized" select="replace(*/@xml:base, '^zip:|[/!]+$', '')"/>
        <xsl:variable name="this-target-base-directory" select="replace($this-target-base-normalized, '^(.+/)([^/]+)$', '$1output/$2' || $output-url-infix || '.components/')"/>
        <xsl:variable name="this-href-pass-1" select="$this-target-base-directory || */@_archive-path"/>
        <xsl:variable name="this-href-pass-2"
            select="replace(replace($this-href-pass-1, $characters-forbidden-in-fullpaths-regex, '_'), ' ', '%20')"
        />
        <xsl:variable name="this-target-subdirectory"
            select="replace($this-href-pass-2, '/[^/]+$', '/')"/>
        <xsl:choose>
            <xsl:when test="$diagnostics-on">
                <!-- In diagnostic mode, you'll see @xml:base unchangned, but @new-href shows where the components will be saved -->
                <xsl:for-each select="*">
                    <xsl:copy>
                        <xsl:copy-of select="@*"/>
                        <xsl:attribute name="new-href" select="$this-href-pass-2"/>
                    </xsl:copy>
                </xsl:for-each>
            </xsl:when>
            <xsl:otherwise>
                <xsl:choose>
                    <xsl:when test="exists(_directory) and not(file:exists($this-target-subdirectory))" use-when="$advanced-functions-available">
                        <xsl:message select="'Creating directory at ' || $this-target-subdirectory"/>
                        <xsl:sequence select="file:create-dir($this-target-subdirectory)"/>
                    </xsl:when>
                    <xsl:when test="exists(_directory)">
                        <xsl:message select="'Empty subdirectory ' || $this-href-pass-2 || ' will be skipped.'"/>
                    </xsl:when>
                    <xsl:when test="exists(_base64Binary)" use-when="$advanced-functions-available">
                        <xsl:if test="not(file:exists($this-target-subdirectory))">
                            <xsl:message select="'Creating directory at ' || $this-target-subdirectory"/>
                            <xsl:sequence select="file:create-dir($this-target-subdirectory)"/>
                        </xsl:if>
                        <xsl:sequence
                            select="file:write-binary($this-href-pass-2, xs:base64Binary(.))"/>
                    </xsl:when>
                    <xsl:when test="exists(_base64Binary)">
                        <xsl:message
                            select="'Item targeted for ' || $this-href-pass-2 || ' is base 64 binary, but advanced functions are not available.'"
                        />
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:result-document href="{$this-href-pass-2}">
                            <xsl:document>
                                <xsl:apply-templates mode="clean-up-archive"/>
                            </xsl:document>
                        </xsl:result-document>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

</xsl:stylesheet>
