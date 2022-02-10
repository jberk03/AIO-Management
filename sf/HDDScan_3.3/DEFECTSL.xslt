<?xml version="1.0" encoding="utf-8"?>

<xsl:stylesheet version="1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

<xsl:template match="/">
  <html xmlns="http://www.w3.org/1999/xhtml" >
    <head>
      <title>HDDScan SCSI Defects Report</title>
    </head>
    <body>
      <br />
      <div style="font-weight: bold; z-index: 102; left: 66px; color: navy;
        font-family: 'Arial'; position: absolute; top: 25px; height: 11px; background-color: white">
		  <span style="width: 500px">
			  HDDScan SCSI Defects Report
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
			<span stype="font-size: 10pt; font-weight: normal">
				Defect list: <xsl:value-of select="HDDScanReport/@list" />
			</span>
			<br />
			<span stype="font-size: 10pt; font-weight: normal">
				Logical format
			</span>
		<br />
			<br />
			<xsl:for-each select="HDDScanReport/DEFECTS">

			<div style="font-size: 10pt;font-weight: bold;  left: 66px; color: navy;
        		font-family: 'Arial'; height: 11px; background-color: white">
		  		Defects Information
  			</div>
			

			<div style="border: solid 1 black">
				<table border="1" cellpadding="2" cellspacing="0" style="width: 100%">
					<tr>
						
						<td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;
                    				height: 16px; background-color: white; font-size: 8pt; text-align: center" width="30%">
											Num
						</td>
						<td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;
                    				height: 16px; background-color: white; font-size: 8pt; text-align: center" width="70%">
											LBA
						</td>
	
					</tr>
				   <xsl:for-each select="Record">
								
							
					<tr>

						<td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;
                    				height: 16px; background-color: white; font-size: 8pt; text-align: Left" width="30%">
											<xsl:value-of select="@Num"/>
						</td>
						<td bgcolor="white" style="font-weight:normal;color: black; font-family: Arial;
                    				height: 16px; background-color: white; font-size: 8pt; text-align: Left" width="70%">
											<xsl:value-of select="@LBA"/>
						</td>

					</tr>

				   </xsl:for-each>
				</table>
			</div>
			<br />

			<br />
			</xsl:for-each>
		</div>

	</body>
  </html>
</xsl:template>

</xsl:stylesheet> 
