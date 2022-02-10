<?xml version="1.0" encoding="utf-8"?>

<xsl:stylesheet version="1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

<xsl:template match="/">
  <html xmlns="http://www.w3.org/1999/xhtml" >
    <head>
      <title>HDDScan S.M.A.R.T. Report</title>
    </head>
    <body>
      <br />
      <div style="font-weight: bold; z-index: 102; left: 66px; color: navy;
        font-family: 'Arial'; position: absolute; top: 25px; height: 11px; background-color: white">
		  <span style="width: 500px">
			  HDDScan S.M.A.R.T. Report
		  </span>
		  
      </div>
		<br />
		<div style="font-weight: normal; z-index: 101; left: 66px; width: 500px; position: absolute;
        top: 88px; font-size: 10pt; font-family: Arial">
			<span stype="font-color: white; font-size: 10pt; font-weight: normal; bgcolor: black">
				Model:  <xsl:value-of select="HDDScanReport/@model" />
			</span>
			<br />
			<span stype="font-color: white; font-size: 10pt; font-weight: normal; bgcolor: black">
				Firmware:  <xsl:value-of select="HDDScanReport/@firmware" />
			</span>
			<br />
			<span stype="font-color: white; font-size: 10pt; font-weight: normal; bgcolor: black">
				Serial:  <xsl:value-of select="HDDScanReport/@serial" />
			</span>
			<br />
			<span stype="font-color: white; font-size: 10pt; font-weight: normal; bgcolor: black">
				LBA:  <xsl:value-of select="HDDScanReport/@LBA" />
			</span>
			<br />
			<br />
			<span stype="font-size: 10pt; font-weight: normal">
				Report By:  <xsl:value-of select="HDDScanReport/@app" />
			</span>
			<br />
			<span stype="font-size: 10pt; font-weight: normal">
				Report Date: <xsl:value-of select="HDDScanReport/@time" />
			</span>
		<br />
			<br />
      <xsl:for-each select="HDDScanReport/SSD">
        <div style="width: 400px">
          <table cellpadding="0" cellspacing="0" style="width: 100%">
            <xsl:for-each select="Feature">
              <tr>
                <td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;height: 12px; background-color: white; font-size: 8pt; text-align: left" width="35%">
                  <xsl:value-of select="@Name"/>
                </td>
                <td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;height: 12px; background-color: white; font-size: 8pt; text-align: left">
                  <xsl:value-of select="@Value"/>
                </td>
              </tr>
            </xsl:for-each>
          </table>
        </div>
        <br />
        <br />  
      </xsl:for-each>
      
      
			<xsl:for-each select="HDDScanReport/SMART">
				<div style="border: solid 1 black">
					<table border="1" cellpadding="2" cellspacing="0" style="width: 100%">
						<tr>
						<td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;
                    height: 16px; background-color: white; font-size: 8pt; text-align: center" width="7%">
											
			</td>
						<td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;
                    height: 16px; background-color: white; font-size: 8pt; text-align: center" width="7%">
											Num
			</td>
						<td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;
                    height: 16px; background-color: white; font-size: 8pt; text-align: center" width="30%">
											Attribute Name
			</td>
						<td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;
                    height: 16px; background-color: white; font-size: 8pt; text-align: center" width="7%">
											Value
			</td>
						<td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;
                    height: 16px; background-color: white; font-size: 8pt; text-align: center" width="7%">
											Worst
			</td>
						<td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;
                    height: 16px; background-color: white; font-size: 8pt; text-align: center" width="40%">
											Raw(hex)
			</td>
						<td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;
                    height: 16px; background-color: white; font-size: 8pt; text-align: center" width="7%">
											Threshold
			</td>
						</tr>
						<xsl:for-each select="Property">
								
							
									<tr>



							<td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;height: 16px; background-color: white; font-size: 8pt; text-align: center" width="7%">
							<xsl:choose>
								<xsl:when test="@Image='Red'">
								<img src="Red.ico"/>						
								</xsl:when>
								<xsl:when test="@Image='Yellow'">
								<img src="Yellow.ico"/>					
								</xsl:when>
								<xsl:otherwise> 
								<img src="Green.ico"/>						
								</xsl:otherwise>
							</xsl:choose>
							<br/>
							</td>
			<td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;
                    height: 16px; background-color: white; font-size: 8pt; text-align: center" width="7%">
											<xsl:value-of select="@Num"/>
			</td>
			<td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;
                    height: 16px; background-color: white; font-size: 8pt; text-align: left" width="43%">
											<xsl:value-of select="@Attribute"/>
			</td>
			<td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;
                    height: 16px; background-color: white; font-size: 8pt; text-align: center" width="7%">
											<xsl:value-of select="@Value"/>
			</td>
			<td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;
                    height: 16px; background-color: white; font-size: 8pt; text-align: center" width="7%">
											<xsl:value-of select="@Worst"/>
			</td>
																				<td bgcolor="white" style="font-weight: normal;color:black; font-family: Arial;
                    height: 16px; background-color: white; font-size: 8pt; text-align: center" width="17%">
											<xsl:value-of select="@Raw"/>
			</td>
			<td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;
                    height: 16px; background-color: white; font-size: 8pt; text-align: center" width="7%">
											<xsl:value-of select="@Threshold" />
			</td>
									</tr>

						</xsl:for-each>
					</table>
				</div>
				<br />
			</xsl:for-each>
		</div>

	</body>
  </html>
</xsl:template>

</xsl:stylesheet> 
