<map version="1.0.1">
<!-- To view this file, download free mind mapping software FreeMind from http://freemind.sourceforge.net -->
<node BACKGROUND_COLOR="#3399ff" CREATED="1408765001143" ID="ID_785005460" MODIFIED="1408984827307" TEXT="Printer Project">
<icon BUILTIN="forward"/>
<node BACKGROUND_COLOR="#66ff66" CREATED="1408765746406" FOLDED="true" ID="ID_1859045130" MODIFIED="1408986683869" POSITION="right" STYLE="bubble" TEXT="Get List Of Targets">
<edge COLOR="#66ff66"/>
<node CREATED="1408765915941" ID="ID_1350919258" MODIFIED="1408984363090" STYLE="fork" TEXT="Save users on each floor as a CSV file"/>
<node CREATED="1408769380775" ID="ID_1606961067" MODIFIED="1408984363090" STYLE="fork" TEXT="If list is usernames, generate their respective computer names"/>
<node BACKGROUND_COLOR="#66ff66" CREATED="1408765946374" FOLDED="true" ID="ID_1836238808" MODIFIED="1408986680960" TEXT="Verify users are online">
<node BACKGROUND_COLOR="#66ff66" CREATED="1408766005319" FOLDED="true" ID="ID_1304631563" MODIFIED="1408984476922" TEXT="If users are not online generate new list to be used later">
<node CREATED="1408767149030" ID="ID_1969474810" MODIFIED="1408984363089" STYLE="fork" TEXT="Include in report"/>
</node>
<node CREATED="1408766024694" ID="ID_1877646450" MODIFIED="1408984363089" STYLE="fork" TEXT="Online users will be added to a target list"/>
</node>
</node>
<node BACKGROUND_COLOR="#00ffcc" CREATED="1408766060838" FOLDED="true" ID="ID_383996687" MODIFIED="1408986689143" POSITION="right" STYLE="bubble" TEXT="Send Payload">
<edge COLOR="#00ffcc"/>
<node CREATED="1408766078390" ID="ID_1527814354" MODIFIED="1408984732832" STYLE="fork" TEXT="Each online machine will be sent to the C: root directory">
<edge COLOR="#00ffcc"/>
</node>
</node>
<node BACKGROUND_COLOR="#33ffff" CREATED="1408766142837" FOLDED="true" ID="ID_688635869" MODIFIED="1408986718535" POSITION="right" STYLE="bubble" TEXT="Execute Payload">
<edge COLOR="#33ffff"/>
<node BACKGROUND_COLOR="#33ffff" CREATED="1408765194444" FOLDED="true" ID="ID_1427661465" MODIFIED="1408986687439" TEXT="Delete Retired Printers">
<edge COLOR="#33ffff"/>
<node BACKGROUND_COLOR="#33ffff" CREATED="1408767724823" FOLDED="true" ID="ID_145306568" MODIFIED="1408986685367" TEXT="Query installed printers stored from registry">
<edge COLOR="#33ffff"/>
<node CREATED="1408768737832" ID="ID_328866329" MODIFIED="1408984644867" STYLE="fork" TEXT="Save to preQuery.txt file">
<edge COLOR="#33ffff"/>
</node>
</node>
<node CREATED="1408767739670" ID="ID_549385427" MODIFIED="1408984644866" STYLE="fork" TEXT="If list contains retired printer, remove it">
<edge COLOR="#33ffff"/>
</node>
</node>
<node BACKGROUND_COLOR="#33ffff" CREATED="1408765249270" FOLDED="true" ID="ID_1667830999" MODIFIED="1408986707508" TEXT="Add New Sharp">
<edge COLOR="#33ffff"/>
<node CREATED="1408767900455" ID="ID_768842031" MODIFIED="1408984644865" STYLE="fork" TEXT="DllImport( &quot;winspool.drv&quot; ) can be used to install printer">
<edge COLOR="#33ffff"/>
</node>
<node CREATED="1408767948486" ID="ID_1601538823" MODIFIED="1408984644863" STYLE="fork" TEXT="public static extern bool AddPrinterConnection(string PrinterName);">
<edge COLOR="#33ffff"/>
</node>
<node BACKGROUND_COLOR="#33ffff" CREATED="1408768758391" FOLDED="true" ID="ID_724862274" MODIFIED="1408986691891" TEXT="Query printers again to check for installation">
<edge COLOR="#33ffff"/>
<node CREATED="1408768781622" ID="ID_202838476" MODIFIED="1408984644854" STYLE="fork" TEXT="Save to postQuery.txt file">
<edge COLOR="#33ffff"/>
</node>
</node>
</node>
<node BACKGROUND_COLOR="#33ffff" CREATED="1408765378405" FOLDED="true" ID="ID_1365796975" MODIFIED="1408986703380" TEXT="Set Default Printer">
<edge COLOR="#33ffff"/>
<node CREATED="1408768044007" ID="ID_1159161947" MODIFIED="1408984644854" STYLE="fork" TEXT="Leave GUI OnLogon to set default printer, use winspool.drv">
<edge COLOR="#33ffff"/>
</node>
<node CREATED="1408768111607" ID="ID_378348741" MODIFIED="1408984644854" STYLE="fork" TEXT="Send directions to change default">
<edge COLOR="#33ffff"/>
</node>
<node BACKGROUND_COLOR="#33ffff" CREATED="1408768139207" FOLDED="true" ID="ID_751009363" MODIFIED="1408986701732" TEXT="Contact HC to set for user">
<edge COLOR="#33ffff"/>
<node CREATED="1408768161111" ID="ID_342887369" MODIFIED="1408984644853" STYLE="fork" TEXT="Implement button in Command Center">
<edge COLOR="#33ffff"/>
</node>
</node>
</node>
</node>
<node BACKGROUND_COLOR="#ff9999" CREATED="1408766157622" FOLDED="true" ID="ID_478366395" MODIFIED="1408986709772" POSITION="right" STYLE="bubble" TEXT="Delete Payload">
<edge COLOR="#ff9999"/>
<node CREATED="1408768250071" ID="ID_297239655" MODIFIED="1408984235927" STYLE="fork" TEXT="Delete from C:"/>
</node>
<node BACKGROUND_COLOR="#ff9999" CREATED="1408765274246" FOLDED="true" ID="ID_1523065612" MODIFIED="1408986717523" POSITION="right" STYLE="bubble" TEXT="Generate Report">
<edge COLOR="#ff9999"/>
<node BACKGROUND_COLOR="#ff9999" CREATED="1408768308519" FOLDED="true" ID="ID_322184130" MODIFIED="1408986712332" TEXT="Main Section">
<edge COLOR="#ff9999"/>
<node CREATED="1408768998536" ID="ID_1597305369" MODIFIED="1408984846659" STYLE="fork" TEXT="Location"/>
<node CREATED="1408769005815" ID="ID_167098937" MODIFIED="1408984235926" STYLE="fork" TEXT="Name"/>
<node CREATED="1408769008471" ID="ID_1605380117" MODIFIED="1408984235926" STYLE="fork" TEXT="Retired printer installed?"/>
<node CREATED="1408769028039" ID="ID_386404454" MODIFIED="1408984235926" STYLE="fork" TEXT="Successful deletion of retired printer?"/>
<node CREATED="1408769042856" ID="ID_615728059" MODIFIED="1408984235925" STYLE="fork" TEXT="Sharp printer installed?"/>
</node>
<node BACKGROUND_COLOR="#ff9999" CREATED="1408768468231" FOLDED="true" ID="ID_398457154" MODIFIED="1408986714020" TEXT="Problems Section">
<edge COLOR="#ff9999"/>
<node CREATED="1408769093592" ID="ID_1084366287" MODIFIED="1408984849379" STYLE="fork" TEXT="Location"/>
<node CREATED="1408769097096" ID="ID_1276902196" MODIFIED="1408984235925" STYLE="fork" TEXT="Name"/>
<node CREATED="1408769106023" ID="ID_652430849" MODIFIED="1408984235924" STYLE="fork" TEXT="Issue"/>
</node>
<node BACKGROUND_COLOR="#ff9999" CREATED="1408768860167" FOLDED="true" ID="ID_995943838" MODIFIED="1408986715340" TEXT="Offline Section">
<edge COLOR="#ff9999"/>
<node CREATED="1408769185239" ID="ID_1537837451" MODIFIED="1408984853507" STYLE="fork" TEXT="Location"/>
<node CREATED="1408769195127" ID="ID_1836634158" MODIFIED="1408984235923" STYLE="fork" TEXT="Name"/>
</node>
</node>
</node>
</map>
