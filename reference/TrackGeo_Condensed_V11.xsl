<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.1" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:inr="http://mycompany.com/mynamespace">
  <xsl:include href="../format_V2.xsl" />
  <xsl:param name="xslRootDirectory" select="inr:xslRootDirectory" />
  <!-- Variable to hold unit string -->
  <xsl:variable name="unit">
    <xsl:choose>
      <xsl:when test="//@linearUnits = 'Imperial'">ft</xsl:when>
      <xsl:otherwise>m</xsl:otherwise>
    </xsl:choose>
  </xsl:variable>
  <!-- Setting Out Table -->
  <xsl:template match="/">
    <xsl:variable name="gridOut" select="inr:SetGridOut(number(InRoads/@outputGridScaleFactor))" />
    <html>
      <head>
        <link rel="stylesheet" type="text/css" href="{$xslRootDirectory}/_Themes/engineer/theme_V2.css" />
        <!-- Title displayed in browser Title Bar -->
        <title lang="en">Setting Out Table</title>
      </head>
      <body>
        <xsl:choose>
          <xsl:when test="$xslShowHelp = 'true'">
            <xsl:call-template name="StyleSheetHelp" />
          </xsl:when>
          <xsl:otherwise>
            <xsl:for-each select="InRoads/GeometryProject/HorizontalAlignment">
              <p lang="en">
                <u>TRACK GEOMETRY DATA -  <xsl:value-of select="@name" /></u>
              </p>
              <table border="1" cellpadding="2" cellspacing="0" width="100%">
                <thead>
                  <tr>
                    <th rowspan="2" valign="middle">ELEMENT</th>
                    <th rowspan="2" valign="middle">CURVE No.</th>
                    <th rowspan="2" valign="middle">POINT</th>
                    <th rowspan="2" valign="middle">STATION</th>
                    <th rowspan="2" valign="middle">BEARING<br /></th>
                    <th colspan="2" valign="middle">COORDINATES</th>
                    <th rowspan="2" colspan="4" valign="middle">DATA</th>
                  </tr>
                  <tr>
                    <th style="font-size: 72%" valign="middle">Northing</th>
                    <th style="font-size: 72%" valign="middle">Easting</th>
                  </tr>
                </thead>
                <tbody>
                  <xsl:for-each select="HorizontalElements">
                    <xsl:apply-templates />
                  </xsl:for-each>
                </tbody>
              </table>
            </xsl:for-each>
          </xsl:otherwise>
        </xsl:choose>
      </body>
    </html>
  </xsl:template>
<!-- Horizontal Linear Data -->
  <xsl:template match="HorizontalLine">
    <tr style="font-size: 72%">
      <td class="sidepad" align="center">TANGENT</td>
      <td> </td>
      <td class="sidepad" align="center">
          <xsl:text>PI</xsl:text>
        </td>
      <td class="sidepad" align="center">
        <xsl:value-of select="inr:stationFormat(number(Start/station/@externalStation), $xslStationFormat, $xslStationPrecision, string(Start/station/@externalStationName))" />
      </td>
      <td class="sidepad" align="center" nowrap="nowrap">
        <xsl:value-of select="inr:directionFormat(number(@direction), $xslDirectionFormat, $xslDirectionPrecision, $xslDirectionModeFormat, $xslAngularMethod)" />
      </td>
      <td class="sidepad" align="center">
        <xsl:value-of select="inr:northingFormat(number(Start/@northing), $xslNorthingPrecision)" />
      </td>
      <td class="sidepad" align="center">
        <xsl:value-of select="inr:eastingFormat(number(Start/@easting), $xslEastingPrecision)" />
      </td>
      <td colspan="4"> </td>
    </tr>
    <xsl:if test="End[@pointType = 'POE']">
      <tr style="font-size: 72%">
        <td class="sidepad" align="center" style="border-top: none ;border-bottom: none">TANGENT</td>
        <td> </td>
        <td class="sidepad" align="center">
          <xsl:text>POL</xsl:text>
        </td>
        <td> </td>
        <td class="sidepad" align="center" nowrap="nowrap">
        <xsl:value-of select="inr:directionFormat(number(@direction), $xslDirectionFormat, $xslDirectionPrecision, $xslDirectionModeFormat, $xslAngularMethod)" />
      </td>
      <td> </td>
      <td> </td>
        <td colspan="4"> </td>
      </tr>
    </xsl:if>
  </xsl:template>
  <!-- Horizontal Circular Data -->
  <xsl:template match="HorizontalCircle">
    <tr style="font-size: 72%">
      <td style="border-bottom: none"> </td>
      <td rowspan="3"> </td>
      <td class="sidepad" align="center">
        <xsl:value-of select="Start/@type" />
      </td>
      <td class="sidepad" align="center">
        <xsl:value-of select="inr:stationFormat(number(Start/station/@externalStation), $xslStationFormat, $xslStationPrecision, string(Start/station/@externalStationName))" />
      </td>
      <td rowspan="3"> </td>
      <td class="sidepad" align="center">
        <xsl:value-of select="inr:northingFormat(number(Start/@northing), $xslEastingPrecision)" />
      </td>
      <td align="center" class="sidepad">
        <xsl:value-of select="inr:eastingFormat(number(Start/@easting), $xslEastingPrecision)" />
      </td>
      <td class="sidepad" align="center" style="border-bottom: none; border-right: none ">c = 
                <xsl:value-of select="inr:angularFormat(number(@delta), $xslAngularFormat, $xslAngularPrecision, $xslAngularMethod)" /></td>
      <td class="sidepad" align="center" style="border-bottom: none ;border-left: none; border-right: none">Dc=
                <xsl:value-of select="inr:angularFormat(number(@degreeOfCurve), $xslAngularFormat, $xslAngularPrecision, $xslAngularMethod)" /></td>
      <td class="sidepad" align="center" style="border-bottom: none ;border-right: none;border-left: none">R=
                <xsl:value-of select="inr:distanceFormat(number(@radius), $xslDistancePrecision)" />'
            </td>
      <td class="sidepad" align="center" style="border-bottom: none ;border-left: none">Lc=
                <xsl:value-of select="inr:distanceFormat(number((End/station/@externalStation)-(Start/station/@externalStation)), $xslDistancePrecision)" />'
            </td>
    </tr>
    <tr style="font-size: 72%">
      <td style="border-top: none;border-bottom: none" align="center">CURVE</td>
      <td class="sidepad" align="center">
        <xsl:value-of select="PI/@type" />
      </td>
      <td> </td>
      


      <xsl:choose>
        <xsl:when test="//CurvesetPoint/@curveSetStartElement &lt;= current()/@elementNumber and //CurvesetPoint/@curveSetStopElement &gt;= current()/@elementNumber">
         <td class="sidepad" align="center">
            <xsl:value-of select="inr:northingFormat(number(//CurvesetPoint[@curveSetStartElement &lt;= current()/@elementNumber][@curveSetStopElement &gt;= current()/@elementNumber]/GeometryPoint/@northing), $xslNorthingPrecision)"/>
          </td>
          <td class="sidepad" align="center">
            <xsl:value-of select="inr:eastingFormat(number(//CurvesetPoint[@curveSetStartElement &lt;= current()/@elementNumber][@curveSetStopElement &gt;= current()/@elementNumber]/GeometryPoint/@easting), $xslEastingPrecision)"/>
          </td>
        </xsl:when>
        <xsl:otherwise><td>&#xa0;</td></xsl:otherwise>
      </xsl:choose>
      
      
<!--      
      <td class="sidepad" align="center">
        <xsl:value-of select="inr:northingFormat(number(PI/@northing), $xslNorthingPrecision)" />
      </td>
      <td class="sidepad" align="center">
        <xsl:value-of select="inr:eastingFormat(number(PI/@easting), $xslEastingPrecision)" />
      </td>
-->      
     
<!--
      <td align="center" style="border-bottom: none ;border-right: none">V=      MPH</td>
      <td align="center" style="border-top: none ;border-bottom: none ;border-left: none; border-right: none">Et=      "</td>
      <td align="center" style="border-top: none ;border-bottom: none ;border-left: none; border-right: none">Ea=      "</td>
      <td align="center" style="border-top: none ;border-bottom: none ;border-left: none">Eu=      "</td>
-->
      <!-- Get CANT values -->
      <xsl:call-template name="GetCantInfo">
        <xsl:with-param name="startIntSta" select="Start/station/@internalStation" />
        <xsl:with-param name="endIntSta" select="End/station/@internalStation" />
      </xsl:call-template>
    </tr>
    <tr style="font-size: 72%">
      <td style="border-top: none"> </td>
      <td class="sidepad" align="center">
        <xsl:value-of select="End/@type" />
      </td>
      <td align="center" class="sidepad">
        <xsl:value-of select="inr:stationFormat(number(End/station/@externalStation), $xslStationFormat, $xslStationPrecision, string(End/station/@externalStationName))" />
      </td>
      <td align="center" class="sidepad">
        <xsl:value-of select="inr:northingFormat(number(End/@northing), $xslNorthingPrecision)" />
      </td>
      <td align="center" class="sidepad">
        <xsl:value-of select="inr:eastingFormat(number(End/@easting), $xslEastingPrecision)" />
      </td>
      <td class="sidepad" align="center" style="border-top: none ; border-right: none; border-right: none">Tc=
                <xsl:value-of select="inr:distanceFormat(number(@tangentLength), $xslDistancePrecision)" />'
            </td>
      <td class="sidepad" align="center" style="border-top: none ; border-left: none; border-right: none">Ec=
                <xsl:value-of select="inr:distanceFormat(number(@externalDistance), $xslDistancePrecision)" />'
            </td>
      <td class="sidepad" align="center" style="border-top: none ; border-right: none; border-left: none">CC:N
                <xsl:value-of select="inr:northingFormat(number(Center/@northing), $xslNorthingPrecision)" /></td>
      <td class="sidepad" align="center" style="border-top: none ; border-left: none">E
                <xsl:value-of select="inr:eastingFormat(number(Center/@easting), $xslEastingPrecision)" /></td>
    </tr>
    <xsl:if test="position() = last()">
      <tr style="font-size: 72%">
        <td> </td>
        <td> </td>
        <td> </td>
        <td class="sidepad" align="center">
          <xsl:value-of select="End/@type" />
        </td>
        <td class="sidepad" align="center" nowrap="nowrap">
          <xsl:value-of select="inr:stationFormat(number(End/station/@externalStation), $xslStationFormat, $xslStationPrecision, string(End/station/@externalStationName))" />
        </td>
        <td class="sidepad" align="center">
          <xsl:value-of select="inr:northingFormat(number(End/@northing), $xslNorthingPrecision)" />
        </td>
        <td class="sidepad" align="center">
          <xsl:value-of select="inr:eastingFormat(number(End/@easting), $xslEastingPrecision)" />
        </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
      </tr>
    </xsl:if>
  </xsl:template>
  <!-- Horizontal Spiral Data -->
  <xsl:template match="HorizontalSpiral">
    <tr style="font-size: 72%">
      <td rowspan="2" valign="middle" align="center">SPIRAL</td>
      <td rowspan="2"> </td>
      <xsl:choose>
        <xsl:when test=".//Start/@type = 'CS'">
         <td> </td>
         <td> </td>
         <td rowspan="2"> </td>
         <td> </td>
         <td> </td>
        </xsl:when>
        <xsl:otherwise>
      <td class="sidepad" align="center">
        <xsl:value-of select="Start/@type" />
      </td>
      <td class="sidepad" align="center" nowrap="nowrap">
        <xsl:value-of select="inr:stationFormat(number(Start/station/@externalStation), $xslStationFormat, $xslStationPrecision, string(Start/station/@externalStationName))" />
      </td>
      <td rowspan="2"> </td>
         <td class="sidepad" align="center">
        <xsl:value-of select="inr:northingFormat(number(Start/@northing), $xslNorthingPrecision)" />
      </td>
      <td class="sidepad" align="center">
        <xsl:value-of select="inr:eastingFormat(number(Start/@easting), $xslEastingPrecision)" />
      </td>
        </xsl:otherwise>
      </xsl:choose>
      <td class="sidepad" align="center" style="border-bottom: none ;border-right: none">s =
                <xsl:value-of select="inr:angularFormat(number(@thetaAngle), $xslAngularFormat, $xslAngularPrecision, $xslAngularMethod)" /></td>
      <td class="sidepad" align="center" style="border-bottom: none ;border-left: none; border-right: none">Ls=
                <xsl:value-of select="inr:distanceFormat(number(@length), $xslDistancePrecision)" />'
            </td>
      <td class="sidepad" align="center" style="border-bottom: none ;border-left: none; border-right: none">LT=
                <xsl:value-of select="inr:distanceFormat(number(@longTangent), $xslDistancePrecision)" />'
            </td>
      <td class="sidepad" align="center" style="border-bottom: none ;border-left: none">STs=
                <xsl:value-of select="inr:distanceFormat(number(@shortTangent), $xslDistancePrecision)" />'
            </td>
    </tr>
    <tr style="font-size: 72%">
      <xsl:choose>
        <xsl:when test=".//End/@type = 'SC'">
         <td> </td>
         <td> </td>
         <td> </td>
         <td> </td>
        </xsl:when>
        <xsl:otherwise>
      <td class="sidepad" align="center">
        <xsl:value-of select="End/@type" />
      </td>
      <td class="sidepad" align="center" nowrap="nowrap">
        <xsl:value-of select="inr:stationFormat(number(End/station/@externalStation), $xslStationFormat, $xslStationPrecision, string(End/station/@externalStationName))" />
      </td>
         <td class="sidepad" align="center">
        <xsl:value-of select="inr:northingFormat(number(End/@northing), $xslNorthingPrecision)" />
      </td>
      <td class="sidepad" align="center">
        <xsl:value-of select="inr:eastingFormat(number(End/@easting), $xslEastingPrecision)" />
      </td>
        </xsl:otherwise>
      </xsl:choose>
      <td class="sidepad" align="center" style="border-top: none ;border-bottom: none ;border-right: none">Xs=
                <xsl:value-of select="inr:distanceFormat(number(@xs), $xslDistancePrecision)" />'
            </td>
      <td class="sidepad" align="center" style="border-top: none ;border-bottom: none ;border-left: none; border-right: none">Ys=
                <xsl:value-of select="inr:distanceFormat(number(@ys), $xslDistancePrecision)" />'
            </td>
      <td class="sidepad" align="center" style="border-top: none ;border-bottom: none ;border-left: none; border-right: none">P=
                <xsl:value-of select="inr:distanceFormat(number(@p), $xslDistancePrecision)" />'
            </td>
      <td class="sidepad" align="center" style="border-top: none ;border-bottom: none ;border-left: none">K=
                <xsl:value-of select="inr:distanceFormat(number(@ks), $xslDistancePrecision)" />'
            </td>
    </tr>
    <xsl:if test="position() = last()">
      <tr style="font-size: 72%">
        <td> </td>
        <td> </td>
        <td> </td>
        <td class="sidepad" align="center">
          <xsl:value-of select="End/@type" />
        </td>
        <td class="sidepad" align="center" nowrap="nowrap">
          <xsl:value-of select="inr:stationFormat(number(End/station/@externalStation), $xslStationFormat, $xslStationPrecision, string(End/station/@externalStationName))" />
        </td>
        <td> </td>
        <td class="sidepad" align="center">
          <xsl:value-of select="inr:northingFormat(number(End/@northing), $xslNorthingPrecision)" />
        </td>
        <td class="sidepad" align="center">
          <xsl:value-of select="inr:eastingFormat(number(End/@easting), $xslEastingPrecision)" />
        </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
        <td> </td>
      </tr>
    </xsl:if>
  </xsl:template>
  <xsl:template name="StyleSheetHelp">
    <div class="section1">
      <h4 lang="en">Notes</h4>
      <p class="normal1" lang="en">
                You must include at least one horizontal alignment in the <em>Include</em> field on the
                <em>Tools &gt; XML Reports &gt; Geometry</em> command to get results from this report.
            </p>
      <p class="small" lang="en">
        <em>© 2006 Bentley Systems, Inc</em>
      </p>
    </div>
  </xsl:template>
  <xsl:template name="GetCantInfo">
    <xsl:param name="startIntSta" />
    <xsl:param name="endIntSta" />
    <xsl:variable name="transType">
      <xsl:text>Linear</xsl:text>
    </xsl:variable>
    <xsl:for-each select="../../Superelevation/CantAlignment/Cant">
      <xsl:if test="$transType = @transitionType">
        <xsl:if test="$startIntSta = Start/station/@internalStation">
          <xsl:if test="$endIntSta = End/station/@internalStation">
            <xsl:variable name="mph">
              <xsl:text>V= </xsl:text>
              <xsl:value-of select="inr:distanceFormat(number(@designSpeed), $xslDistancePrecision)" />
              <xsl:text> MPH</xsl:text>
            </xsl:variable>
            <xsl:variable name="et">
              <xsl:text>Et= </xsl:text>
              <xsl:value-of select="inr:distanceFormat(number(@equilibriumCant), $xslDistancePrecision)" />
            </xsl:variable>
            <xsl:variable name="ea">
              <xsl:text>Ea= </xsl:text>
              <xsl:value-of select="inr:distanceFormat(number(@appliedCant), $xslDistancePrecision)" />
            </xsl:variable>
            <xsl:variable name="eu">
              <xsl:text>Eu= </xsl:text>
              <xsl:value-of select="inr:distanceFormat(number(@cantDeficiency), $xslDistancePrecision)" />
            </xsl:variable>
            <td align="center" style="border-top: none ;border-bottom: none ;border-right: none">
              <xsl:value-of select="$mph" />
            </td>
            <td align="center" style="border-top: none ;border-bottom: none ;border-left: none; border-right: none">
              <xsl:value-of select="$et" />
            </td>
            <td align="center" style="border-top: none ;border-bottom: none ;border-left: none; border-right: none">
              <xsl:value-of select="$ea" />
            </td>
            <td align="center" style="border-top: none ;border-bottom: none ;border-left: none">
              <xsl:value-of select="$eu" />
            </td>
          </xsl:if>
        </xsl:if>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>
</xsl:stylesheet>