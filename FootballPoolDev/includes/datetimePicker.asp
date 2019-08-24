	<!-- Pop-up calendar for date selection. -->
	<div id="calendar">
		<form action="#" method="get" onsubmit="return false;">
			<table id="calendarTable">
				<tr class="monthHeader">
					<th><a href="#" title="Show previous month." onclick="return monthClick(-1);">&lt;</a></th>
					<th id="calendarMonth" colspan="5">&nbsp;</th>
					<th><a href="#" title="Show next month." onclick="return monthClick(1);">&gt;</a></th>
				</tr>
				<tr class="daysHeader">
					<th>Su</th>
					<th>Mo</th>
					<th>Tu</th>
					<th>We</th>
					<th>Th</th>
					<th>Fr</th>
					<th>Sa</th>
				</tr>
				<tr class="dates">
					<td class="weekend"><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
			  		<td class="weekend"><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
				</tr>
				<tr class="dates">
					<td class="weekend"><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td class="weekend"><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
				</tr>
				<tr class="dates">
					<td class="weekend"><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
				<td class="weekend"><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
				</tr>
				<tr class="dates">
					<td class="weekend"><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td class="weekend"><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
				</tr>
				<tr class="dates">
					<td class="weekend"><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td class="weekend"><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
				</tr>
				<tr class="dates">
				<td class="weekend"><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
					<td class="weekend"><a href="#" onclick="return dateClick(this);">&nbsp;</a></td>
				</tr>
			</table>
			<p><input type="submit" name="submit" value=" Ok " class="button" title="Use the selected date." onclick="return okCalendarClick();" />&nbsp;<input type="submit" name="submit" value="Cancel" class="button" title="Cancel the change." onclick="return cancelCalendarClick();" /></p>
		</form>
	</div>

	<!-- Pop-up clock for time selection. -->
	<div id="clock">
		<form id="clockForm" action="#" method="get" onsubmit="return false;">
			<table>
				<tr>
					<td><select name="clockHour">
						<option value="1">1</option>
						<option value="2">2</option>
						<option value="3">3</option>
						<option value="4">4</option>
						<option value="5">5</option>
						<option value="6">6</option>
						<option value="7">7</option>
						<option value="8">8</option>
						<option value="9">9</option>
						<option value="10">10</option>
						<option value="11">11</option>
						<option value="12">12</option>
					</select><strong>:</strong><select name="clockMinute">
						<option value="00">00</option>
						<option value="05">05</option>
						<option value="10">10</option>
						<option value="15">15</option>
						<option value="20">20</option>
						<option value="25">25</option>
						<option value="30">30</option>
						<option value="35">35</option>
						<option value="40">40</option>
						<option value="45">45</option>
						<option value="50">50</option>
						<option value="55">55</option>
					</select>&nbsp;<select name="clockMeridiem">
						<option value="am">am</option>
						<option value="pm">pm</option>
					</select></td>
				</tr>
			</table>
			<p><input type="submit" name="submit" value=" Ok " class="button" title="Use the selected time." onclick="return okClockClick();" />&nbsp;<input type="submit" name="submit" value="Cancel" class="button" title="Cancel the change." onclick="return cancelClockClick();" /></p>
		</form>
	</div>
