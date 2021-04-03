<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
   <xsl:output method="text" indent="yes"/>
   <xsl:template match="node()">
      <xsl:value-of select="." />
   </xsl:template>
</xsl:stylesheet>