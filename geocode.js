/*
Use google maps to return Lat & Lng values based on an address.
By M. Evanoff
*/

var addressCol = 8,
	latCol = 12,
	lngCol = 13;

function getLatLng() {

	var ss = SpreadsheetApp.getActiveSpreadsheet(),
		sheet = ss.getActiveSheet(),
		startingRow = emptyStart_(sheet, latCol),
		addrRange = sheet.getRange(startingRow, addressCol, sheet.getMaxRows(), 1),
		addrData = addrRange.getValues();

	for (var i = 0; i < addrData.length; ++i) {
		var address = addrData[i],
			response = Maps.newGeocoder().geocode(address);

		var lat = response.results[0].geometry.location.lat, //latitude as returned by Google Maps
			lng = response.results[0].geometry.location.lng, //longitude as returned by Google Maps
			cell = (i + startingRow);

		sheet.getRange(cell, latCol).setValue(lat);
		sheet.getRange(cell, lngCol).setValue(lng);

		//Timeout so we don't blow past the time-based geocoding limitations
		Utilities.sleep(1000);

	}
}

//Only run the geocode on rows that don't have a latitude

function emptyStart_(sheet, latCol) {
	var latColData = sheet.getRange(2, latCol, sheet.getMaxRows(), 1).getValues();
	for (var i = 0; i < latColData.length; ++i) {
		latInfo = latColData[i];
		if (latInfo == "") {
			return i + 2;
		}
	}
}