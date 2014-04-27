GSpreadsheetDev
===============
<h3>Larimer County Parcel Consumer</h3>
<p>Little console application that take a list of parcels, pings the Larimer County parcel API, and updates a Google Spreasheet. 
</p>

<h4>Usage</h4>
<p>Update appKeys in App.Config with your Google account info and set the target Goole Spreadsheet. Run or debug app in Visual Studio or MonoDevelop.</p>
<pre><code>
&lt;!-- google account --&gt;
&lt;add key="gAccount" value=""/&gt;

&lt;!-- google password --&gt;
&lt;add key="gPassword" value=""/&gt;

&lt;!-- name of the servce, what ever you call it --&gt;
&lt;add key="sheetService" value="FGPService"/&gt;

&lt;!-- the name of the worksheet --&gt;
&lt;add key="sheetTitle" value="FHGPParcels"/&gt;

&lt;!-- target worksheet, zero based, 0 is first worksheet in the workbook, 1 is the 2nd worksheet --&gt;
&lt;add key="worksheetNumber" value="2"/&gt;

&lt;!-- column that contains parcel numbers, zero-based, 0 is Column A, 2 is Column B --&gt;
&lt;add key="parcelColumnNumber" value="1"/&gt;
</code></pre>

<h4>Worksheet</h4>
<p>App expects worksheet to have the following column names:</p>
<ul>
<li>Parcel Number</li>
<li>Owner</li>	
<li>Bulding</li>	
<li>Occupancy</li> 
<li>Description</li>	
<li>Address	City</li>	
<li>Zip</li>	
<li>Actual Value</li>	
<li>Serial Number</li>	
<li>Legal	</li>
<li>Subdivision</li>	
<li>Filing</li>	
<li>IsMember</li>	
<li>AnnexMarkerColor</li>	
<li>PolygonColor</li>	
<li>AddressCombo</li>	
<li>GrantSigned</li>	
<li>GrantFiled</li>	
<li>MemberId</li>	
<li>MailAddress1</li>	
<li>MailAddress2</li>	
<li>MailAddressCity</li>	
<li>MailAddressState</li>	
<li>MailAddressZip</li>	
<li>Owner1</li>	
<li>Owner2</li>
</ul>
