<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
	<link rel="stylesheet" href="styles/newstyles/reset.css" type="text/css" media="screen" title="no title" />
	<link rel="stylesheet" href="styles/newstyles/text.css" type="text/css" media="screen" title="no title" />
	<link rel="stylesheet" href="styles/newstyles/form.css" type="text/css" media="screen" title="no title" />
	<link rel="stylesheet" href="styles/newstyles/buttons.css" type="text/css" media="screen" title="no title" />
	<link rel="stylesheet" href="styles/newstyles/grid.css" type="text/css" media="screen" title="no title" />	
	<link rel="stylesheet" href="styles/newstyles/layout.css" type="text/css" media="screen" title="no title" />	
	<link rel="stylesheet" href="styles/newstyles/ui-darkness/jquery-ui-1.8.12.custom.css" type="text/css" media="screen" title="no title" />
	<link rel="stylesheet" href="styles/newstyles/plugin/jquery.visualize.css" type="text/css" media="screen" title="no title" />
	<link rel="stylesheet" href="styles/newstyles/plugin/facebox.css" type="text/css" media="screen" title="no title" />
	<link rel="stylesheet" href="styles/newstyles/plugin/uniform.default.css" type="text/css" media="screen" title="no title" />
	<link rel="stylesheet" href="styles/newstyles/plugin/dataTables.css" type="text/css" media="screen" title="no title" />
	<link rel="stylesheet" href="styles/newstyles/custom.css" type="text/css" media="screen" title="no title">
    
    <script type="text/javascript" src="scripts/common.js"></script>
	<script type="text/javascript" src="scripts/menu.js"></script>
	<script type="text/javascript" src="scripts/tableSort.js"></script>

    
    <script src="js/jquery/jquery-1.5.2.min.js"></script>
	<script src="js/jquery/jquery-ui-1.8.12.custom.min.js"></script>
<!--    <script src="js/misc/excanvas.min.js"></script>-->
    <script src="js/jquery/facebox.js"></script>
    <script src="js/jquery/jquery.visualize.js"></script>
    <script src="js/jquery/jquery.dataTables.min.js"></script>
    <script src="js/jquery/jquery.tablesorter.min.js"></script>
    <script src="js/jquery/jquery.uniform.min.js"></script>
    <script src="js/jquery/jquery.placeholder.min.js"></script>
    
<!--    <script src="js/widgets.js"></script>-->
<!--    <script src="js/dashboard.js"></script>-->
    
    <script type="text/javascript" charset="utf-8">
			$(document).ready(function() {
				$('#example').dataTable( {
					"aaSorting": [[ 4, "desc" ]]
				} );
			} );
	</script>
    
<title>Untitled Document</title>
</head>

<body>






<div id="wrapper">
        
        <div id="top">
            
            <div class="content_pad">			
                <ul class="right">
                    <li><a href="javascript:;" class="top_icon"><span class="ui-icon ui-icon-person"></span>Logged in as John Doe</a></li>
                    <li><a href="javascript:;" class="new_messages top_alert">1 New Message</a></li>
                    <li><a href="../pages/settings.html">Settings</a></li>
                    <li><a href="../index.html">Logout</a></li>
                </ul>
            </div> <!-- .content_pad -->
            
        </div> <!-- #top -->	
        
        <div id="header">
            
            <div class="content_pad">
                <h1><a href="../dashboard.html">Dashboard Admin</a></h1>
                
                <ul id="nav">
                    <li class="nav_icon"><a href="../dashboard.html"><span class="ui-icon ui-icon-home"></span>Home</a></li>
    
                    <li class="nav_dropdown nav_current nav_icon">
                        <a href="javascript:;"><span class="ui-icon ui-icon-gripsmall-diagonal-se"></span>Styles</a>
                        
                        <div class="nav_menu">			
                            <ul>
                                <li><a href="../pages/text.html">Buttons &amp; Text</a></li>	
                                <li><a href="../pages/grid.html">Grid Layout</a></li>	
                                <li><a href="../pages/tables.html">Tables</a></li>	
                                <li><a href="../pages/forms.html">Forms</a></li>	
                                <li><a href="../pages/charts.html">Charts</a></li>						
                            </ul>
                            
                        </div>
                    </li>
                    
                    <li class="nav_icon"><a href="../pages/widgets.html"><span class="ui-icon ui-icon-gear"></span>Widgets</a></li>
                    <li class="nav_icon"><a href="../pages/reports.html"><span class="ui-icon ui-icon-signal"></span>Reports</a></li>
                    
                    <li class="nav_dropdown nav_icon_only">
                        <a href="javascript:;">&nbsp;</a>
                        
                        <div class="nav_menu">
                            
                            <ul>
                                <li><a href="javascript:;">Overflow Menu</a></li>
                                <li><a href="javascript:;">Items Can</a></li>
                                <li><a href="javascript:;">Go Here</a></li>
                            </ul>
                        </div> <!-- .menu -->
                    </li>
                </ul>
            </div> <!-- .content_pad -->
            
        </div> <!-- #header -->	
        
        <div id="masthead">
            
            <div class="content_pad">
                
                <h1 class="">Table Styles</h1>
                
                <div id="bread_crumbs">
                    <a href="../dashboard.html">Home</a> / 
                    <a href="javascript:;" class="current_page">Table Styles</a>				
                 </div> <!-- #bread_crumbs -->
                
                <div id="search">
                    <form action="/search" method="get">
                        <input type="text" value="" placeholder="Search" name="search" id="search_input" title="Search" />					
                        <input type="submit" value="" name="submit" class="submit" />					
                    </form>
                </div> <!-- #search -->
                
            </div> <!-- .content_pad -->
            
        </div> <!-- #masthead -->	
        
        
        
        
            <div id="content" class="xgrid">
            
            <div class="x12">
                   
                        <h2>Data Table Plugin</h2>

                    
                    <table class="data display datatable" id="example">
                        <thead>
                            <tr>
                                <th>Rendering engine</th>
                                <th>Browser</th>
                                <th>Platform(s)</th>
                                <th>Engine version</th>
                                <th>CSS grade</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr class="odd gradeX">
                                <td>Trident</td>
                                <td>Internet
                                     Explorer 4.0</td>
                                <td>Win 95+</td>
                                <td class="center"> 4</td>
                                <td class="center">X</td>
                            </tr>
                            <tr class="even gradeC">
                                <td>Trident</td>
                                <td>Internet
                                     Explorer 5.0</td>
                                <td>Win 95+</td>
                                <td class="center">5</td>
                                <td class="center">C</td>
                            </tr>
                            <tr class="odd gradeA">
                                <td>Trident</td>
                                <td>Internet
                                     Explorer 5.5</td>
                                <td>Win 95+</td>
                                <td class="center">5.5</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="even gradeA">
                                <td>Trident</td>
                                <td>Internet
                                     Explorer 6</td>
                                <td>Win 98+</td>
                                <td class="center">6</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="odd gradeA">
                                <td>Trident</td>
                                <td>Internet Explorer 7</td>
                                <td>Win XP SP2+</td>
                                <td class="center">7</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="even gradeA">
                                <td>Trident</td>
                                <td>AOL browser (AOL desktop)</td>
                                <td>Win XP</td>
                                <td class="center">6</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="gradeA">
                                <td>Gecko</td>
                                <td>Firefox 1.0</td>
                                <td>Win 98+ / OSX.2+</td>
                                <td class="center">1.7</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="gradeA">
                                <td>Gecko</td>
                                <td>Firefox 1.5</td>
                                <td>Win 98+ / OSX.2+</td>
                                <td class="center">1.8</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="gradeA">
                                <td>Gecko</td>
                                <td>Firefox 2.0</td>
                                <td>Win 98+ / OSX.2+</td>
                                <td class="center">1.8</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="gradeA">
                                <td>Gecko</td>
                                <td>Firefox 3.0</td>
                                <td>Win 2k+ / OSX.3+</td>
                                <td class="center">1.9</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="gradeA">
                                <td>Gecko</td>
                                <td>Camino 1.0</td>
                                <td>OSX.2+</td>
                                <td class="center">1.8</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="gradeA">
                                <td>Gecko</td>
                                <td>Camino 1.5</td>
                                <td>OSX.3+</td>
                                <td class="center">1.8</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="gradeA">
                                <td>Gecko</td>
                                <td>Netscape 7.2</td>
                                <td>Win 95+ / Mac OS 8.6-9.2</td>
                                <td class="center">1.7</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="gradeA">
                                <td>Gecko</td>
                                <td>Netscape Browser 8</td>
                                <td>Win 98SE+</td>
                                <td class="center">1.7</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="gradeA">
                                <td>Gecko</td>
                                <td>Netscape Navigator 9</td>
                                <td>Win 98+ / OSX.2+</td>
                                <td class="center">1.8</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="gradeA">
                                <td>Gecko</td>
                                <td>Mozilla 1.0</td>
                                <td>Win 95+ / OSX.1+</td>
                                <td class="center">1</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="gradeA">
                                <td>Gecko</td>
                                <td>Mozilla 1.1</td>
                                <td>Win 95+ / OSX.1+</td>
                                <td class="center">1.1</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="gradeA">
                                <td>Gecko</td>
                                <td>Mozilla 1.2</td>
                                <td>Win 95+ / OSX.1+</td>
                                <td class="center">1.2</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="gradeA">
                                <td>Gecko</td>
                                <td>Mozilla 1.3</td>
                                <td>Win 95+ / OSX.1+</td>
                                <td class="center">1.3</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="gradeA">
                                <td>Gecko</td>
                                <td>Mozilla 1.4</td>
                                <td>Win 95+ / OSX.1+</td>
                                <td class="center">1.4</td>
                                <td class="center">A</td>
                            </tr>
                            <tr class="gradeA">
                                <td>Gecko</td>
							<td>Mozilla 1.5</td>
							<td>Win 95+ / OSX.1+</td>
							<td class="center">1.5</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Gecko</td>
							<td>Mozilla 1.6</td>
							<td>Win 95+ / OSX.1+</td>
							<td class="center">1.6</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Gecko</td>
							<td>Mozilla 1.7</td>
							<td>Win 98+ / OSX.1+</td>
							<td class="center">1.7</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Gecko</td>
							<td>Mozilla 1.8</td>
							<td>Win 98+ / OSX.1+</td>
							<td class="center">1.8</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Gecko</td>
							<td>Seamonkey 1.1</td>
							<td>Win 98+ / OSX.2+</td>
							<td class="center">1.8</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Gecko</td>
							<td>Epiphany 2.20</td>
							<td>Gnome</td>
							<td class="center">1.8</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Webkit</td>
							<td>Safari 1.2</td>
							<td>OSX.3</td>
							<td class="center">125.5</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Webkit</td>
							<td>Safari 1.3</td>
							<td>OSX.3</td>
							<td class="center">312.8</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Webkit</td>
							<td>Safari 2.0</td>
							<td>OSX.4+</td>
							<td class="center">419.3</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Webkit</td>
							<td>Safari 3.0</td>
							<td>OSX.4+</td>
							<td class="center">522.1</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Webkit</td>
							<td>OmniWeb 5.5</td>
							<td>OSX.4+</td>
							<td class="center">420</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Webkit</td>
							<td>iPod Touch / iPhone</td>
							<td>iPod</td>
							<td class="center">420.1</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Webkit</td>
							<td>S60</td>
							<td>S60</td>
							<td class="center">413</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Presto</td>
							<td>Opera 7.0</td>
							<td>Win 95+ / OSX.1+</td>
							<td class="center">-</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Presto</td>
							<td>Opera 7.5</td>
							<td>Win 95+ / OSX.2+</td>
							<td class="center">-</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Presto</td>
							<td>Opera 8.0</td>
							<td>Win 95+ / OSX.2+</td>
							<td class="center">-</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Presto</td>
							<td>Opera 8.5</td>
							<td>Win 95+ / OSX.2+</td>
							<td class="center">-</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Presto</td>
							<td>Opera 9.0</td>
							<td>Win 95+ / OSX.3+</td>
							<td class="center">-</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Presto</td>
							<td>Opera 9.2</td>
							<td>Win 88+ / OSX.3+</td>
							<td class="center">-</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Presto</td>
							<td>Opera 9.5</td>
							<td>Win 88+ / OSX.3+</td>
							<td class="center">-</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Presto</td>
							<td>Opera for Wii</td>
							<td>Wii</td>
							<td class="center">-</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Presto</td>
							<td>Nokia N800</td>
							<td>N800</td>
							<td class="center">-</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>Presto</td>
							<td>Nintendo DS browser</td>
							<td>Nintendo DS</td>
							<td class="center">8.5</td>
							<td class="center">C/A<sup>1</sup></td>
						</tr>
						<tr class="gradeC">
							<td>KHTML</td>
							<td>Konqureror 3.1</td>
							<td>KDE 3.1</td>
							<td class="center">3.1</td>
							<td class="center">C</td>
						</tr>
						<tr class="gradeA">
							<td>KHTML</td>
							<td>Konqureror 3.3</td>
							<td>KDE 3.3</td>
							<td class="center">3.3</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeA">
							<td>KHTML</td>
							<td>Konqureror 3.5</td>
							<td>KDE 3.5</td>
							<td class="center">3.5</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeX">
							<td>Tasman</td>
							<td>Internet Explorer 4.5</td>
							<td>Mac OS 8-9</td>
							<td class="center">-</td>
							<td class="center">X</td>
						</tr>
						<tr class="gradeC">
							<td>Tasman</td>
							<td>Internet Explorer 5.1</td>
							<td>Mac OS 7.6-9</td>
							<td class="center">1</td>
							<td class="center">C</td>
						</tr>
						<tr class="gradeC">
							<td>Tasman</td>
							<td>Internet Explorer 5.2</td>
							<td>Mac OS 8-X</td>
							<td class="center">1</td>
							<td class="center">C</td>
						</tr>
						<tr class="gradeA">
							<td>Misc</td>
							<td>NetFront 3.1</td>
							<td>Embedded devices</td>
							<td class="center">-</td>
							<td class="center">C</td>
						</tr>
						<tr class="gradeA">
							<td>Misc</td>
							<td>NetFront 3.4</td>
							<td>Embedded devices</td>
							<td class="center">-</td>
							<td class="center">A</td>
						</tr>
						<tr class="gradeX">
							<td>Misc</td>
							<td>Dillo 0.8</td>
							<td>Embedded devices</td>
							<td class="center">-</td>
							<td class="center">X</td>
						</tr>
						<tr class="gradeX">
							<td>Misc</td>
							<td>Links</td>
							<td>Text only</td>
							<td class="center">-</td>
							<td class="center">X</td>
						</tr>
						<tr class="gradeX">
							<td>Misc</td>
							<td>Lynx</td>
							<td>Text only</td>
							<td class="center">-</td>
							<td class="center">X</td>
						</tr>
						<tr class="gradeC">
							<td>Misc</td>
							<td>IE Mobile</td>
							<td>Windows Mobile 6</td>
							<td class="center">-</td>
							<td class="center">C</td>
						</tr>
						<tr class="gradeC">
							<td>Misc</td>
							<td>PSP browser</td>
							<td>PSP</td>
							<td class="center">-</td>
							<td class="center">C</td>
						</tr>
						<tr class="gradeU">
							<td>Other browsers</td>
							<td>All others</td>
							<td>-</td>
							<td class="center">-</td>
							<td class="center">U</td>
						</tr>
					</tbody>
				</table>
			
		</div> <!-- .x12 -->
		
	</div> <!-- #content -->
    
    
    

 


</body>
</html>
