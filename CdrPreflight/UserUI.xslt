<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:frmwrk="Corel Framework Data" exclude-result-prefixes="frmwrk">
  <xsl:output method="xml" encoding="UTF-8" indent="yes"/>

  <frmwrk:uiconfig>
    <frmwrk:compositeNode xPath="/uiConfig/commandBars/commandBarData[@guid='f3016f3c-2847-4557-b61a-a2d05319cf18']"/>
    <frmwrk:compositeNode xPath="/uiConfig/frame"/>
  </frmwrk:uiconfig>

  <xsl:template match="node()|@*">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="uiConfig/commandBars/commandBarData[@guid='f3016f3c-2847-4557-b61a-a2d05319cf18']/menubar/modeData[@guid='76d73481-9076-44c9-821c-52de9408cce2']/item[@guidRef='6c91d5ab-d648-4364-96fb-3e71bcfaf70d']">
    <xsl:copy-of select="."/>
    <item guidRef="9387393f-8a16-4ee5-9ef5-ef9f4f8eb5b9"/>
  </xsl:template>
  
  <xsl:template match="dockSheet[@guidRef='6884106d-f37e-4712-986d-b2fe7e31ecdf']">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
      <dockPage guidRef="34695a15-b045-1b43-96a4-6e5eee9679c7"/>
    </xsl:copy>
  </xsl:template>
  
</xsl:stylesheet>