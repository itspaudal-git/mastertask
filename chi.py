from __future__ import division

import datetime
import json
import math
import re
import sys
from optparse import OptionParser
import urllib3

http = urllib3.PoolManager(
    cert_reqs='CERT_REQUIRED',
    ca_certs='/path/to/your/cabundle.pem'
)


import numpy as np
import tovala_utilities.meal_parsing.plating_guide_parsing as pg_parse
import urllib3 as url
from tovala_utilities import labor_planning
from tovala_utilities.labor_planning import equipment as eq
from tovala_utilities.labor_planning import facilities, scheduler
from tovala_utilities.labor_planning import tasks as TKS
from tovala_utilities.labor_planning.constants import *

import google_sheets as gs
from ignore import get_sheet_data_as_list

# Define Google Sheets details for ignorelists
workstation_id = '1mSVlGySk-GZL4oToRUpmYYp8_QwaSOk814I0woJdx0U'
range_name = 'ignore_production_planner!A1:A400'  
id_sheet_data = 'sheet_id!B1'
id_sheet_data_range = get_sheet_data_as_list(workstation_id,id_sheet_data)
ignore_list = get_sheet_data_as_list(workstation_id, range_name)

N_MEALS_PER_BOX_AVG = 5.6

VL_INFORMATIONAL = 1
VL_DEBUG = 2
VL_WARNING = 3
VL_ERROR = 4

#________________________________________________________________
#THIS VERSION COMPATIBLE WITH tovala_utilities v 0.4.0 and Higher

WORK_CENTER_CATEGORY_DISPLAY = {"Hot Line": gs.COLOR_SALMON, "Mixing": gs.COLOR_LIGHT_GREEN, "Prepping": gs.COLOR_YELLOW, "Portioning and Sealing": gs.COLOR_LIGHT_BLUE}

PROD_DAYS = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday2","Monday2"]

FACILITY_HOURS = {"Pershing": {"min": 0.0, "mid": 12.0, "max": 20.0}, "Tubeway": {"min": 0.0, "mid": 10.0, "max": 12.0}}

TIME_STUDY_WORK_CENTERS = ["Tray Sealers Pershing", "Manual Portioning Pershing", "Sleeving Pershing"]

COLOR_CORNFLOWER_BLUE = {"red": 74.0/256, "blue": 232.0/256, "green": 134.0/256}
COLOR_DARK_YELLOW_2 = {"red": 191.0/256, "blue": 0.0/256, "green": 144.0/256}
COLOR_CYAN = {"red": 0.0/256, "blue": 255.0/256, "green": 255.0/256}
COLOR_GRADIENT_GREEN = {"red": 217.0/256, "blue": 211.0/256, "green": 234.0/256}
COLOR_GRADIENT_YELLOW = {"red": 255.0/256, "blue": 102.0/256, "green": 214.0/256}
COLOR_GRADIENT_RED = {"red": 230.0/256, "blue": 115.0/256, "green": 124.0/256}


def compileConditionalFormattingBgGradient(sheetId, requests, startEndRowColumn, service, spreadsheet_id, minpoint, midpoint, maxpoint, colorMin = COLOR_GRADIENT_GREEN, colorMid = COLOR_GRADIENT_YELLOW, colorMax = COLOR_GRADIENT_RED): 

    requests.append({
        'addConditionalFormatRule': {
            'rule': {
                'ranges': [{
                    'sheetId': sheetId,
                    'startRowIndex': startEndRowColumn[0],
                    'endRowIndex': startEndRowColumn[1],
                    'startColumnIndex': startEndRowColumn[2],
                    'endColumnIndex': startEndRowColumn[3]
                }],
                'gradientRule': {
                    'minpoint': {
                        'type': "NUMBER",
                        'value': str(minpoint),
                        'color': colorMin
                    },
                    'midpoint': {
                        'type': "NUMBER",
                        'value': str(midpoint),
                        'color': colorMid
                    },
                    'maxpoint': {
                        'type': "NUMBER",
                        'value': str(maxpoint),
                        'color': colorMax
                    }
                }
            },
            "index": 0
        }
    })

    return requests

def compileFormatCells_HoursCountFormat(sheetId, requests, startEndRowColumn, service, spreadsheet_id): 

    requests.append({
        'repeatCell': {
            'range': {
                'sheetId': sheetId,
                'startRowIndex': startEndRowColumn[0],
                'endRowIndex': startEndRowColumn[1],
                'startColumnIndex': startEndRowColumn[2],
                'endColumnIndex': startEndRowColumn[3]
            },
            'cell': {
                'userEnteredFormat': {
                    'numberFormat': {
                        'type': "NUMBER",
                        'pattern': "0.0 \"h\""
                    }
                }
            },
            'fields': 'userEnteredFormat(numberFormat)'
        }
    })

    return requests

def generate_permission_groups(admin_users, staffing):
	tovala_production_planning_users = [user for user in admin_users]
	tovala_production_planning_users.extend([user['email'] for user in staffing['corporate']])
	tovala_production_planning_users.extend([user['email'] for user in staffing['plant_managers']])

	tovala_production_mod_and_greater_users = [user for user in tovala_production_planning_users]
	tovala_production_mod_and_greater_users.extend([user['email'] for user in staffing['mods']])

	tovala_production_execution_users = [user for user in tovala_production_mod_and_greater_users]
	tovala_production_execution_users.extend([user['email'] for user in staffing['supervisors']])

	return [tovala_production_planning_users,tovala_production_mod_and_greater_users,tovala_production_execution_users]

def generate_fn_task_compatibility(facility_network, _sheetColor, _spreadsheetInfo, _requests, _data, _service, _spreadsheet_id, _sheetIDs, _admin_users):
	[_requests, fnTaskCompatibilitySheetId] = gs.compileAddSheetRequest(facility_network.name.capitalize() + " Task Compatibility", _sheetColor, _spreadsheetInfo, _requests, _service, _spreadsheet_id, sheetID = _sheetIDs['fn_task_compability'])

	lenTaskCompatibilityList = len(TKS.validTaskTypes)+4

		#Protect and Hide Sheet
	try:
		sheet = next(sheet for sheet in _spreadsheetInfo['sheets'] if sheet['properties']['sheetId'] == fnTaskCompatibilitySheetId)
		_requests = gs.compileProtectSheet(sheet, _requests, _service, _spreadsheet_id, gs.A1NotationToRC("A1"), _admin_users)
		tabRows = sheet['properties']['gridProperties']['rowCount']
		tabColumns = sheet['properties']['gridProperties']['columnCount']
	except StopIteration:
		_requests = gs.compileProtectSheet("", _requests, _service, _spreadsheet_id, gs.A1NotationToRC("A1"), _admin_users, newSheet = {"newSheet": True, "id": fnTaskCompatibilitySheetId})
		tabRows = 1000
		tabColumns = 26
	_requests = gs.compileHideSheet(fnTaskCompatibilitySheetId, _requests)
	_requests = gs.compileFormatCells_TextFormat(fnTaskCompatibilitySheetId, _requests, gs.A1NotationToRC("A1:" + gs.columnIndexToLetter(lenTaskCompatibilityList) + "1"), _service, _spreadsheet_id, textFormat = {'bold': True, 'fontFamily': 'Calibri', 'fontSize': 11})
	_requests = gs.compileBatchColorCellRangeColorV1(fnTaskCompatibilitySheetId, _requests, gs.A1NotationToRC("A1:" + gs.columnIndexToLetter(lenTaskCompatibilityList) + "1"), _service, _spreadsheet_id, bgColor = gs.COLOR_DARK_GRAY_2)
	_requests = gs.compileFormatCells_TextFormat(fnTaskCompatibilitySheetId, _requests, gs.A1NotationToRC("A2:A100"), _service, _spreadsheet_id, textFormat = {'bold': True, 'fontFamily': 'Calibri', 'fontSize': 11})
	_requests = gs.compileBatchColorCellRangeColorV1(fnTaskCompatibilitySheetId, _requests, gs.A1NotationToRC("A2:A100"), _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_GRAY)
	_requests = gs.compileBordersBox(fnTaskCompatibilitySheetId, _requests, gs.A1NotationToRC("A1:" + gs.columnIndexToLetter(lenTaskCompatibilityList) + "1"), _service, _spreadsheet_id)
	if(tabColumns < lenTaskCompatibilityList + 3):
		_requests = gs.compileAddDimensionToSheet(fnTaskCompatibilitySheetId, _requests, _service, _spreadsheet_id, dimension = "COLUMNS", countToAppend = (lenTaskCompatibilityList + 3 -tabColumns))

	vals = ["Task Type:"]

	for facility in enumerate(facility_network.facilityList):
		vals.extend([facility[1].name.capitalize() + " Compatible Work Centers:", "", "", ""])
		work_centers = []
		for taskType in TKS.validTaskTypes:
			work_centers.append([work_center.name for work_center in facility[1].work_center_list if taskType in work_center.compatible_process_list])

		_data.append({
				"range": facility_network.name.capitalize() + ' Task Compatibility!B' + str(4*facility[0]+2) + ':' + gs.columnIndexToLetter(len(work_centers)+4) + str(4*facility[0]+5),
				"values": work_centers,
				"majorDimension": "COLUMNS"
			})
		
	_data.append({
			"range": facility_network.name.capitalize() + ' Task Compatibility!A1:A' + str(4*(len(facility_network.facilityList))+1),
			"values": [vals],
			"majorDimension": "COLUMNS"
		})

	_data.append({
			"range": facility_network.name.capitalize() + ' Task Compatibility!B1:'  + gs.columnIndexToLetter(lenTaskCompatibilityList) + '1',
			"values": [TKS.validTaskTypes],
			"majorDimension": "ROWS"
		})

	return [_requests, _data, lenTaskCompatibilityList]

def generate_facility_work_center_lists(facility_network, _sheetColor, _spreadsheetInfo, _requests, _data, _service, _spreadsheet_id, _sheetIDs, _admin_users):
	plannedFacilityWCSheetID = _sheetIDs['starting_facility_work_center']

	maxWorkCenterListColumns = 26
	for facility in facility_network.facilityList:
		if(maxWorkCenterListColumns < 2*len(facility.work_center_list)+4):
			maxWorkCenterListColumns = 2*len(facility.work_center_list)+4

	for facility in facility_network.facilityList:
		[_requests, actualFacilityWCSheetID] = gs.compileAddSheetRequest(facility.name.capitalize() + " Work Centers", _sheetColor, _spreadsheetInfo, _requests, _service, _spreadsheet_id, sheetID = plannedFacilityWCSheetID)
		plannedFacilityWCSheetID = str(int(plannedFacilityWCSheetID)+10000)
		
		try:
			sheet = next(sheet for sheet in _spreadsheetInfo['sheets'] if sheet['properties']['sheetId'] == actualFacilityWCSheetID)
			_requests = gs.compileProtectSheet(sheet, _requests, _service, _spreadsheet_id, gs.A1NotationToRC("A1"), _admin_users)
			tabRows = sheet['properties']['gridProperties']['rowCount']
			tabColumns = sheet['properties']['gridProperties']['columnCount']
		except StopIteration:
			_requests = gs.compileProtectSheet("", _requests, _service, _spreadsheet_id, gs.A1NotationToRC("A1"), _admin_users, newSheet = {"newSheet": True, "id": actualFacilityWCSheetID})
			tabRows = 1000
			tabColumns = 26
		_requests = gs.compileHideSheet(actualFacilityWCSheetID, _requests)
		_requests = gs.compileFormatCells_TextFormat(actualFacilityWCSheetID, _requests, gs.A1NotationToRC("A1:" + gs.columnIndexToLetter(maxWorkCenterListColumns) + "1"), _service, _spreadsheet_id, textFormat = {'bold': True, 'fontFamily': 'Calibri', 'fontSize': 11})
		_requests = gs.compileBatchColorCellRangeColorV1(actualFacilityWCSheetID, _requests, gs.A1NotationToRC("A1:" + gs.columnIndexToLetter(maxWorkCenterListColumns) + "1"), _service, _spreadsheet_id, bgColor = gs.COLOR_DARK_GRAY_2)
		_requests = gs.compileFormatCells_TextFormat(actualFacilityWCSheetID, _requests, gs.A1NotationToRC("A2:A100"), _service, _spreadsheet_id, textFormat = {'bold': True, 'fontFamily': 'Calibri', 'fontSize': 11})
		_requests = gs.compileBatchColorCellRangeColorV1(actualFacilityWCSheetID, _requests, gs.A1NotationToRC("A2:A100"), _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_GRAY)
		_requests = gs.compileBordersBox(actualFacilityWCSheetID, _requests, gs.A1NotationToRC("A1:" + gs.columnIndexToLetter(maxWorkCenterListColumns) + "1"), _service, _spreadsheet_id)
		if(tabColumns < maxWorkCenterListColumns+3):
			_requests = gs.compileAddDimensionToSheet(actualFacilityWCSheetID, _requests, _service, _spreadsheet_id, dimension = "COLUMNS", countToAppend = (maxWorkCenterListColumns + 3 - tabColumns))

		vals = []
		for work_center in facility.work_center_list:
			vals_equip = [work_center.name]
			vals_equip_type = [""]
			for equipment in work_center.workcenter_equipment_list:
				vals_equip.append(equipment.name)
				vals_equip_type.append(equipment.type)
			vals.append(vals_equip)
			vals.append(vals_equip_type)

		_data.append({
				"range": facility.name.capitalize() + ' Work Centers!A1:A2',
				"values": [["Work Centers:", "Equipment Assigned to WC:"]],
				"majorDimension": "COLUMNS"
			})

		_data.append({
				"range": facility.name.capitalize() + ' Work Centers!B1:' + gs.columnIndexToLetter(maxWorkCenterListColumns) + '20',
				"values": vals,
				"majorDimension": "COLUMNS"
			})

	return [_requests, _data, maxWorkCenterListColumns]

def generate_time_study_sheet(_work_center, _facility, _spreadsheetInfo, _requests, _data, _service, _spreadsheet_id, _plannedFacilityWorkCenterTimeStudySheetID):

	[_requests, actualFacilityWorkCenterTimeStudySheetID] = gs.compileAddSheetRequest(_facility.name + ": " + _work_center.name + " Time Studies", gs.COLOR_FOREST_GREEN, _spreadsheetInfo, _requests, _service, _spreadsheet_id, sheetID = _plannedFacilityWorkCenterTimeStudySheetID)
	_plannedFacilityWorkCenterTimeStudySheetID = str(int(_plannedFacilityWorkCenterTimeStudySheetID) + 1)

	_requests = gs.compileFormatCells_TextFormat(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("A1"), _service, _spreadsheet_id, textFormat = {'bold': True, 'fontFamily': 'Arial', 'fontSize': 18})
	_requests = gs.compileFormatCells_TextFormat(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("A3:X3"), _service, _spreadsheet_id, textFormat = {'bold': True, 'fontFamily': 'Arial', 'fontSize': 10})
	_requests = gs.compileFormatCells_TextFormat(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("A4:X4"), _service, _spreadsheet_id, textFormat = {'bold': False, 'italic':True, 'fontFamily': 'Arial', 'fontSize': 10})
	_requests = gs.compileFormatCells_TextFormat(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("A5:X100"), _service, _spreadsheet_id, textFormat = {'bold': False, 'fontFamily': 'Arial', 'fontSize': 12})
	_requests = gs.compileBatchColorCellRangeColorV1(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("A3:X4"), _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_GRAY)
	_requests = gs.compileBordersAll(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("A3:X100"), _service, _spreadsheet_id)
	_requests = gs.compileBordersBox(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("A1:E1"), _service, _spreadsheet_id)
	_requests = gs.compileFormatCells_HorizontalAlign(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("A3:D100"), _service, _spreadsheet_id, horizAlign = "CENTER")
	_requests = gs.compileFormatCells_HorizontalAlign(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("E3:E100"), _service, _spreadsheet_id, horizAlign = "LEFT")
	_requests = gs.compileFormatCells_HorizontalAlign(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("N3:X100"), _service, _spreadsheet_id, horizAlign = "CENTER")
	_requests = gs.compileFormatCells_VerticalAlign(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("A5:X100"), _service, _spreadsheet_id, vertAlign = "MIDDLE")
	_requests = gs.compileHideColumn(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("D1"), _service, _spreadsheet_id, hide = True)
	_requests = gs.compileHideColumn(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("F1:M1"), _service, _spreadsheet_id, hide = True)
	_requests = gs.compileWrapStrategy(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("A3:X100"), _service, _spreadsheet_id, wrapStrategy = "WRAP")
	_requests = gs.compileChangeColumnWidth(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("A1:C1"), _service, _spreadsheet_id, pixelWidth = 50)
	_requests = gs.compileChangeColumnWidth(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("E1"), _service, _spreadsheet_id, pixelWidth = 350)
	_requests = gs.compileChangeColumnWidth(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("N1"), _service, _spreadsheet_id, pixelWidth = 120)
	_requests = gs.compileChangeColumnWidth(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("O1"), _service, _spreadsheet_id, pixelWidth = 50)
	_requests = gs.compileChangeColumnWidth(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("P1"), _service, _spreadsheet_id, pixelWidth = 75)
	_requests = gs.compileChangeColumnWidth(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("Q1:T1"), _service, _spreadsheet_id, pixelWidth = 100)
	_requests = gs.compileChangeColumnWidth(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("U1"), _service, _spreadsheet_id, pixelWidth = 120)
	_requests = gs.compileChangeColumnWidth(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("V1"), _service, _spreadsheet_id, pixelWidth = 175)
	_requests = gs.compileChangeColumnWidth(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("W1"), _service, _spreadsheet_id, pixelWidth = 300)
	_requests = gs.compileChangeColumnWidth(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("X1"), _service, _spreadsheet_id, pixelWidth = 100)
	_requests = gs.compileChangeRowHeight(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("A5:A100"), _service, _spreadsheet_id, pixelHeight = 70)
	_requests = gs.compileRepeatCellRequest(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("A6:A100"), _service, _spreadsheet_id, "=if(isblank(B6),\"\",'Master Task List'!$B$1)")
	_requests = gs.compileRepeatCellRequest(actualFacilityWorkCenterTimeStudySheetID, _requests, gs.A1NotationToRC("U6:U100"), _service, _spreadsheet_id, "=if(isblank(B6),\"\",\"Break | Lunch\")")

	_data.append({
			"range": _facility.name + ': ' + _work_center.name + ' Time Studies!A1',
			"values": [[_work_center.name + " Tasks"]],
			"majorDimension": "COLUMNS"
		})
				
	if _work_center.name == 'Sleeving Pershing':
		_data.append({
				"range": _facility.name + ': ' + _work_center.name + ' Time Studies!A3:X5',
				"values": [["Term", "Cycle", "Meal #", "", "Work Order", "", "", "", "", "", "", "", "", "Target Quantity", "UOM", "Produced Quantity", "Time Setup Start", "Time Start Sleeving", "Time End Sleeving", "Time Cleanup End", "Break", "Tables qty", "Staff", "MoD"],\
					["Período", "Ciclo", "Menú", "", "Orden de Trabajo", "", "", "", "", "", "", "", "", "Cantidad Requerida", "UdM", "Cantidad Producida", "Hora del Comienzo de Preparación", "Hora del Comienzo de Empacado", "Hora del Final de Empacado", "Hora del Final de Limpieza del area", "Descanso (marcar con un círculo)", "Cantidad de Mesas", "Empleados", "Mánager"],\
					["Ex", "1", "0", "", "Example 1", "", "", "", "", "", "", "", "", "1000", "Meals", "1002", "9:00 AM", "9:15 AM", "10:30 AM", "10:45 AM", "Break | Lunch", "2 tables", "John, Mickey, Cal", "Doc"]],
				"majorDimension": "ROWS"
			})

	else:
		_data.append({
				"range": _facility.name + ': ' + _work_center.name + ' Time Studies!A3:X5',
				"values": [["Term", "Cycle", "Meal #", "", "Work Order", "", "", "", "", "", "", "", "", "Target Quantity", "UOM", "Produced Quantity", "Time Setup Start", "Time Start Portioning", "Time End Portioning", "Time Cleanup End", "Break", "Tables qty", "Staff", "MoD"],\
					["Período", "Ciclo", "Menú", "", "Orden de Trabajo", "", "", "", "", "", "", "", "", "Cantidad Requerida", "UdM", "Cantidad Producida", "Hora del Comienzo de Preparación", "Hora del Comienzo de Porcionado", "Hora del Final de Porcionado", "Hora del Final de Limpieza del area", "Descanso (marcar con un círculo)", "Cantidad de Mesas", "Empleados", "Mánager"],\
					["Ex", "1", "0", "", "Example 1", "", "", "", "", "", "", "", "", "2000", "1 oz cups", "2053", "9:00 AM", "9:15 AM", "10:30 AM", "10:45 AM", "Break | Lunch", "Sealing Line 1 + Sealing Line 2", "John, Mickey, Cal", "Doc"]],
				"majorDimension": "ROWS"
			})			

	_data.append({
			"range": _facility.name + ': ' + _work_center.name + ' Time Studies!B6',
			"values": [["=filter('Master Task List'!A6:N900,'Master Task List'!L6:L900=\"" + _work_center.name + "\")"]],
			"majorDimension": "COLUMNS"
		})
	return [_requests, _data, _plannedFacilityWorkCenterTimeStudySheetID]

def generate_equipment_sheet(_equipment, _work_center, _facility, _facility_network, _spreadsheetInfo, _requests, _data, _service, _spreadsheet_id, _admin_users, _plannedEquipmentSheetID, _masterSheetColumns):
	[_requests, actualEquipmentSheetID] = gs.compileAddSheetRequest(_facility.name + ": " + _equipment.name, gs.COLOR_LIGHT_RED, _spreadsheetInfo, _requests, _service, _spreadsheet_id, sheetID = _plannedEquipmentSheetID)
	_plannedEquipmentSheetID = str(int(_plannedEquipmentSheetID) + 1)

	try:
		sheet = next(sheet for sheet in _spreadsheetInfo['sheets'] if sheet['properties']['sheetId'] == actualEquipmentSheetID)
		_requests = gs.compileProtectSheet(sheet, _requests, _service, _spreadsheet_id, gs.A1NotationToRC("A1"), _admin_users)
		tabRows = sheet['properties']['gridProperties']['rowCount']
		tabColumns = sheet['properties']['gridProperties']['columnCount']
	except StopIteration:
		_requests = gs.compileProtectSheet("", _requests, _service, _spreadsheet_id, gs.A1NotationToRC("A1"), _admin_users, newSheet = {"newSheet": True, "id": actualEquipmentSheetID})
		tabRows = 1000
		tabColumns = 26
	if(tabColumns < len(_facility_network.production_days)*10 + 3):
		_requests = gs.compileAddDimensionToSheet(actualEquipmentSheetID, _requests, _service, _spreadsheet_id, dimension = "COLUMNS", countToAppend = (len(_facility_network.production_days)*10 + 3 - tabColumns))
	_requests = gs.compileHideSheet(actualEquipmentSheetID, _requests)

	vals_eq_sheet_0 = []
	vals_eq_sheet_1 = []
	vals_eq_sheet_2 = []
	vals_eq_sheet_3 = []
	vals_eq_sheet_4 = [[],[],[],[]]

	for prod_day in enumerate(_facility_network.production_days):
		_requests = gs.compileRepeatCellRequest(actualEquipmentSheetID, _requests, [22, 102, 5*prod_day[0]+2, 5*prod_day[0]+4], _service, _spreadsheet_id, "=\"--\"")
		vals_eq_sheet_0.extend([prod_day[1], "", "", "", ""])
		vals_eq_sheet_1.extend([prod_day[1] + " Task Order", "Priority", "Man Hours", "# Ppl Assigned", "Estimated Hours"])
		vals_eq_sheet_2.extend(["=ifna(filter('Master Task List'!" + _masterSheetColumns["composite"] + "6:" + _masterSheetColumns["composite"] + ",('Master Task List'!" + _masterSheetColumns["equipment"][0] + "6:" + _masterSheetColumns["equipment"][0] + "=\"" + _equipment.name + "\")*('Master Task List'!" + _masterSheetColumns["day_allocated"] + "6:" + _masterSheetColumns['day_allocated'] + "=\"" + prod_day[1] + "\")),\"\")",\
			"=ifna(filter('Master Task List'!" + _masterSheetColumns['priority'] + "6:" + _masterSheetColumns['priority'] + ",('Master Task List'!" + _masterSheetColumns["equipment"][0] + "6:" + _masterSheetColumns["equipment"][0] + "=\"" + _equipment.name + "\")*('Master Task List'!" + _masterSheetColumns["day_allocated"] + "6:" + _masterSheetColumns['day_allocated'] + "=\"" + prod_day[1] + "\")),\"\")",\
			"=ifna(filter('Master Task List'!" + _masterSheetColumns['manHours'] + "6:" + _masterSheetColumns['estHours'] + ",('Master Task List'!" + _masterSheetColumns["equipment"][0] + "6:" + _masterSheetColumns["equipment"][0] + "=\"" + _equipment.name + "\")*('Master Task List'!" + _masterSheetColumns["day_allocated"] + "6:" + _masterSheetColumns['day_allocated'] + "=\"" + prod_day[1] + "\")),\"\")","",""])
		vals_eq_sheet_3.extend(["=ifna(sort(filter(" + gs.columnIndexToLetter(5*prod_day[0]+1) + "2:" + gs.columnIndexToLetter(5*prod_day[0]+5) + "100, " + gs.columnIndexToLetter(5*prod_day[0]+1) + "2:" + gs.columnIndexToLetter(5*prod_day[0]+1) + "100 <> \"\"),2,TRUE),\"\")", "", "", "", ""])
		vals_eq_sheet_4[0].extend(["=ifna(filter('Master Task List'!" + _masterSheetColumns["composite"] + "6:" + _masterSheetColumns["composite"] + ",('Master Task List'!" + _masterSheetColumns["equipment"][1] + "6:" + _masterSheetColumns["equipment"][1] + "=\"" + _equipment.name + "\")*('Master Task List'!" + _masterSheetColumns["day_allocated"] + "6:" + _masterSheetColumns['day_allocated'] + "=\"" + prod_day[1] + "\")),\"\")",\
			"=ifna(filter('Master Task List'!" + _masterSheetColumns['priority'] + "6:" + _masterSheetColumns['priority'] + ",('Master Task List'!" + _masterSheetColumns["equipment"][1] + "6:" + _masterSheetColumns["equipment"][1] + "=\"" + _equipment.name + "\")*('Master Task List'!" + _masterSheetColumns["day_allocated"] + "6:" + _masterSheetColumns['day_allocated'] + "=\"" + prod_day[1] + "\")),\"\")",\
			"=\"--\"","=\"--\"","=ifna(filter('Master Task List'!" + _masterSheetColumns['estHours'] + "6:" + _masterSheetColumns['estHours'] + ",('Master Task List'!" + _masterSheetColumns["equipment"][1] + "6:" + _masterSheetColumns["equipment"][1] + "=\"" + _equipment.name + "\")*('Master Task List'!" + _masterSheetColumns["day_allocated"] + "6:" + _masterSheetColumns['day_allocated'] + "=\"" + prod_day[1] + "\")),\"\")"])
		vals_eq_sheet_4[1].extend(["=ifna(filter('Master Task List'!" + _masterSheetColumns["composite"] + "6:" + _masterSheetColumns["composite"] + ",('Master Task List'!" + _masterSheetColumns["equipment"][2] + "6:" + _masterSheetColumns["equipment"][2] + "=\"" + _equipment.name + "\")*('Master Task List'!" + _masterSheetColumns["day_allocated"] + "6:" + _masterSheetColumns['day_allocated'] + "=\"" + prod_day[1] + "\")),\"\")",\
			"=ifna(filter('Master Task List'!" + _masterSheetColumns['priority'] + "6:" + _masterSheetColumns['priority'] + ",('Master Task List'!" + _masterSheetColumns["equipment"][2] + "6:" + _masterSheetColumns["equipment"][2] + "=\"" + _equipment.name + "\")*('Master Task List'!" + _masterSheetColumns["day_allocated"] + "6:" + _masterSheetColumns['day_allocated'] + "=\"" + prod_day[1] + "\")),\"\")",\
			"=\"--\"","=\"--\"","=ifna(filter('Master Task List'!" + _masterSheetColumns['estHours'] + "6:" + _masterSheetColumns['estHours'] + ",('Master Task List'!" + _masterSheetColumns["equipment"][2] + "6:" + _masterSheetColumns["equipment"][2] + "=\"" + _equipment.name + "\")*('Master Task List'!" + _masterSheetColumns["day_allocated"] + "6:" + _masterSheetColumns['day_allocated'] + "=\"" + prod_day[1] + "\")),\"\")"])
		vals_eq_sheet_4[2].extend(["=ifna(filter('Master Task List'!" + _masterSheetColumns["composite"] + "6:" + _masterSheetColumns["composite"] + ",('Master Task List'!" + _masterSheetColumns["equipment"][3] + "6:" + _masterSheetColumns["equipment"][3] + "=\"" + _equipment.name + "\")*('Master Task List'!" + _masterSheetColumns["day_allocated"] + "6:" + _masterSheetColumns['day_allocated'] + "=\"" + prod_day[1] + "\")),\"\")",\
			"=ifna(filter('Master Task List'!" + _masterSheetColumns['priority'] + "6:" + _masterSheetColumns['priority'] + ",('Master Task List'!" + _masterSheetColumns["equipment"][3] + "6:" + _masterSheetColumns["equipment"][3] + "=\"" + _equipment.name + "\")*('Master Task List'!" + _masterSheetColumns["day_allocated"] + "6:" + _masterSheetColumns['day_allocated'] + "=\"" + prod_day[1] + "\")),\"\")",\
			"=\"--\"","=\"--\"","=ifna(filter('Master Task List'!" + _masterSheetColumns['estHours'] + "6:" + _masterSheetColumns['estHours'] + ",('Master Task List'!" + _masterSheetColumns["equipment"][3] + "6:" + _masterSheetColumns["equipment"][3] + "=\"" + _equipment.name + "\")*('Master Task List'!" + _masterSheetColumns["day_allocated"] + "6:" + _masterSheetColumns['day_allocated'] + "=\"" + prod_day[1] + "\")),\"\")"])
		vals_eq_sheet_4[3].extend(["=ifna(filter('Master Task List'!" + _masterSheetColumns["composite"] + "6:" + _masterSheetColumns["composite"] + ",('Master Task List'!" + _masterSheetColumns["equipment"][4] + "6:" + _masterSheetColumns["equipment"][4] + "=\"" + _equipment.name + "\")*('Master Task List'!" + _masterSheetColumns["day_allocated"] + "6:" + _masterSheetColumns['day_allocated'] + "=\"" + prod_day[1] + "\")),\"\")",\
			"=ifna(filter('Master Task List'!" + _masterSheetColumns['priority'] + "6:" + _masterSheetColumns['priority'] + ",('Master Task List'!" + _masterSheetColumns["equipment"][4] + "6:" + _masterSheetColumns["equipment"][4] + "=\"" + _equipment.name + "\")*('Master Task List'!" + _masterSheetColumns["day_allocated"] + "6:" + _masterSheetColumns['day_allocated'] + "=\"" + prod_day[1] + "\")),\"\")",\
			"=\"--\"","=\"--\"","=ifna(filter('Master Task List'!" + _masterSheetColumns['estHours'] + "6:" + _masterSheetColumns['estHours'] + ",('Master Task List'!" + _masterSheetColumns["equipment"][4] + "6:" + _masterSheetColumns["equipment"][4] + "=\"" + _equipment.name + "\")*('Master Task List'!" + _masterSheetColumns["day_allocated"] + "6:" + _masterSheetColumns['day_allocated'] + "=\"" + prod_day[1] + "\")),\"\")"])

	vals_eq_sheet_0.extend(vals_eq_sheet_1)
	vals_eq_sheet_2.extend(vals_eq_sheet_3)

	_data.append({
			"range": _facility.name + ": " + _equipment.name + '!A1:' + gs.columnIndexToLetter(10*len(_facility_network.production_days)) + '2',
			"values": [vals_eq_sheet_0, vals_eq_sheet_2],
			"majorDimension": "ROWS"
		})
	_data.append({
			"range": _facility.name + ": " + _equipment.name + '!A22:' + gs.columnIndexToLetter(5*len(_facility_network.production_days)) + '22',
			"values": [vals_eq_sheet_4[0]],
			"majorDimension": "ROWS"
		})
	_data.append({
			"range": _facility.name + ": " + _equipment.name + '!A42:' + gs.columnIndexToLetter(5*len(_facility_network.production_days)) + '42',
			"values": [vals_eq_sheet_4[1]],
			"majorDimension": "ROWS"
		})
	_data.append({
			"range": _facility.name + ": " + _equipment.name + '!A62:' + gs.columnIndexToLetter(5*len(_facility_network.production_days)) + '62',
			"values": [vals_eq_sheet_4[2]],
			"majorDimension": "ROWS"
		})
	_data.append({
			"range": _facility.name + ": " + _equipment.name + '!A82:' + gs.columnIndexToLetter(5*len(_facility_network.production_days)) + '82',
			"values": [vals_eq_sheet_4[3]],
			"majorDimension": "ROWS"
		})

	return [_requests, _data, _plannedEquipmentSheetID]

def generate_facility_labor_hours_tab(_facility, _facility_network, _spreadsheetInfo, _requests, _data, _service, _spreadsheet_id, _admin_users, _plannedLaborSheetID, _masterSheetColumns):

	[_requests, actualLaborSheetID] = gs.compileAddSheetRequest(_facility.name.capitalize() + ": Labor", COLOR_DARK_YELLOW_2, _spreadsheetInfo, _requests, _service, _spreadsheet_id, sheetID = _plannedLaborSheetID)

	try:
		sheet = next(sheet for sheet in _spreadsheetInfo['sheets'] if sheet['properties']['sheetId'] == actualLaborSheetID)
		_requests = gs.compileProtectSheet(sheet, _requests, _service, _spreadsheet_id, gs.A1NotationToRC("A1"), _admin_users)
		tabRows = sheet['properties']['gridProperties']['rowCount']
		tabColumns = sheet['properties']['gridProperties']['columnCount']
	except StopIteration:
		_requests = gs.compileProtectSheet("", _requests, _service, _spreadsheet_id, gs.A1NotationToRC("A1"), _admin_users, newSheet = {"newSheet": True, "id": actualLaborSheetID})
		tabRows = 1000
		tabColumns = 26
	
	work_centers = [work_center.name for work_center in _facility.work_center_list]

	date_headers = [["", "Work Center"]]
	for prod_day in enumerate(_facility_network.production_days):
		date_headers.append([prod_day[1],"=(Date(2017,1,1)+(24*" + str(prod_day[0]) + "+(\'Master Task List\'!B1-1)*168)*3600)/86400+date(2017,1,1)"])
	date_headers.append(["","Total"])

	_requests = gs.compileBatchColorCellRangeColorV1(actualLaborSheetID, _requests, gs.A1NotationToRC("A3"), _service, _spreadsheet_id, bgColor = COLOR_CYAN)
	_requests = gs.compileBatchColorCellRangeColorV1(actualLaborSheetID, _requests, gs.A1NotationToRC("B2:" + gs.columnIndexToLetter(len(date_headers)) + "3"), _service, _spreadsheet_id, bgColor = COLOR_CYAN)

	_requests = gs.compileBordersBox(actualLaborSheetID, _requests, gs.A1NotationToRC("B2:" + gs.columnIndexToLetter(len(date_headers)) + "2"), _service, _spreadsheet_id)
	_requests = gs.compileBordersBox(actualLaborSheetID, _requests, gs.A1NotationToRC("A3:" + gs.columnIndexToLetter(len(date_headers)) + "3"), _service, _spreadsheet_id)
	_requests = gs.compileFormatCells_NumberFormat(actualLaborSheetID, _requests, gs.A1NotationToRC("B3:" + gs.columnIndexToLetter(len(date_headers)-1) + "3"), _service, _spreadsheet_id, numFormat = "DATE")

	_requests = gs.compileChangeColumnWidth(actualLaborSheetID, _requests, gs.A1NotationToRC("A1"), _service, _spreadsheet_id, pixelWidth = 230)
	_requests = gs.compileFormatCells_TextFormat(actualLaborSheetID, _requests, gs.A1NotationToRC("A1"), _service, _spreadsheet_id, textFormat = {'bold': True, 'fontFamily': 'Arial', 'fontSize': 14})

	_requests = gs.compileBordersBox(actualLaborSheetID, _requests, gs.A1NotationToRC("B4:" + gs.columnIndexToLetter(len(date_headers)-1) + str(len(work_centers)+3)), _service, _spreadsheet_id)
	_requests = gs.compileBordersBox(actualLaborSheetID, _requests, gs.A1NotationToRC("A4:" + gs.columnIndexToLetter(len(date_headers)) + str(len(work_centers)+3)), _service, _spreadsheet_id)
	_requests = gs.compileRepeatCellRequest(actualLaborSheetID, _requests, gs.A1NotationToRC("B4:" + gs.columnIndexToLetter(len(date_headers)-1) + str(len(work_centers)+3)), _service, _spreadsheet_id, "=sumifs(\'Master Task List\'!$AR$6:$AR$900,\'Master Task List\'!$G$6:$G$900,\"" + _facility.name.capitalize() + "\",\'Master Task List\'!$L$6:$L$900,$A4,\'Master Task List\'!$AV$6:$AV$900,B$3)")
	_requests = gs.compileRepeatCellRequest(actualLaborSheetID, _requests, gs.A1NotationToRC(gs.columnIndexToLetter(len(date_headers)) + "4:" + gs.columnIndexToLetter(len(date_headers)) + str(len(work_centers)+3)), _service, _spreadsheet_id, "=sum(B4:" + gs.columnIndexToLetter(len(date_headers)-1) + "4)")
	_requests = gs.compileWrapStrategy(actualLaborSheetID, _requests, gs.A1NotationToRC(gs.columnIndexToLetter(len(date_headers)-1) + str(len(work_centers)+4)), _service, _spreadsheet_id, wrapStrategy = "WRAP")

	_requests = gs.compileBatchColorCellRangeColorV1(actualLaborSheetID, _requests, gs.A1NotationToRC("A" + str(len(work_centers)+4) + ":" + gs.columnIndexToLetter(len(date_headers)-2) + str(len(work_centers)+4)), _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_GRAY)
	_requests = gs.compileRepeatCellRequest(actualLaborSheetID, _requests, gs.A1NotationToRC("B" + str(len(work_centers)+4) + ":" + gs.columnIndexToLetter(len(date_headers)-2) + str(len(work_centers)+4)), _service, _spreadsheet_id, "=sumifs(B4:B" + str(len(work_centers)+3) + ",$A4:$A" + str(len(work_centers)+3) + ",\"<>Facility\")")
	_requests = compileFormatCells_HoursCountFormat(actualLaborSheetID, _requests, gs.A1NotationToRC("B4:Z100"), _service, _spreadsheet_id)



	_data.append({
			"range": _facility.name.capitalize() + ': Labor!A1',
			"values": [["Labor Hours " + _facility.name.capitalize()]],
			"majorDimension": "ROWS"
		})

	_data.append({
			"range": _facility.name.capitalize() + ': Labor!A2:' + gs.columnIndexToLetter(len(date_headers)) + '3',
			"values": date_headers,
			"majorDimension": "COLUMNS"
		})

	_data.append({
			"range": _facility.name.capitalize() + ': Labor!A4:A' + str(len(work_centers)+3),
			"values": [work_centers],
			"majorDimension": "COLUMNS"
		})

	_data.append({
			"range": _facility.name.capitalize() + ': Labor!A' + str(len(work_centers)+4),
			"values": [["Non-Facility Hours"]],
			"majorDimension": "ROWS"
		})

	_data.append({
			"range": _facility.name.capitalize() + ': Labor!' + gs.columnIndexToLetter(len(date_headers)-1) + str(len(work_centers)+4) + ':' + gs.columnIndexToLetter(len(date_headers)) + str(len(work_centers)+4),
			"values": [["Total Expected Hours:", "=sum(" + gs.columnIndexToLetter(len(date_headers)) + "4:" + gs.columnIndexToLetter(len(date_headers)) + str(len(work_centers)+3) + ")"]],
			"majorDimension": "ROWS"
		})


	return [_requests, _data]

def generate_facility_sheets(_facility, facility_network, _spreadsheetInfo, _requests, _data, _service, _spreadsheet_id, _admin_users, _plannedFacilityEqSheetID, _plannedFacilityWorkCenterTimeStudySheetID, _masterSheetColumns):

	[_requests, actualFacilityEqSheetID] = gs.compileAddSheetRequest(_facility.name.capitalize() + " Equipment List", gs.COLOR_SALMON, _spreadsheetInfo, _requests, _service, _spreadsheet_id, sheetID = _plannedFacilityEqSheetID)
	plannedLaborSheetID = str(int(_plannedFacilityEqSheetID) + 500)
	[_requests, _data] = generate_facility_labor_hours_tab(_facility, facility_network, _spreadsheetInfo, _requests, _data, _service, _spreadsheet_id, _admin_users, plannedLaborSheetID, _masterSheetColumns)
	_plannedFacilityEqSheetID = str(int(_plannedFacilityEqSheetID)+10000)
		
	try:
		sheet = next(sheet for sheet in _spreadsheetInfo['sheets'] if sheet['properties']['sheetId'] == actualFacilityEqSheetID)
		_requests = gs.compileProtectSheet(sheet, _requests, _service, _spreadsheet_id, gs.A1NotationToRC("A1"), _admin_users)
		tabRows = sheet['properties']['gridProperties']['rowCount']
		tabColumns = sheet['properties']['gridProperties']['columnCount']
	except StopIteration:
		_requests = gs.compileProtectSheet("", _requests, _service, _spreadsheet_id, gs.A1NotationToRC("A1"), _admin_users, newSheet = {"newSheet": True, "id": actualFacilityEqSheetID})
		tabRows = 1000
		tabColumns = 26
	#_requests = gs.compileHideSheet(actualFacilityEqSheetID, _requests)
	eqListHeaderColumnEnd = gs.columnIndexToLetter(len(facility_network.production_days)+2)

	_requests = gs.compileFormatCells_TextFormat(actualFacilityEqSheetID, _requests, gs.A1NotationToRC("A1:"+ eqListHeaderColumnEnd +"1"), _service, _spreadsheet_id, textFormat = {'bold': True, 'fontFamily': 'Calibri', 'fontSize': 11})
	_requests = gs.compileBatchColorCellRangeColorV1(actualFacilityEqSheetID, _requests, gs.A1NotationToRC("A1:"+ eqListHeaderColumnEnd +"1"), _service, _spreadsheet_id, bgColor = gs.COLOR_DARK_GRAY_2)
	_requests = gs.compileBordersBox(actualFacilityEqSheetID, _requests, gs.A1NotationToRC("A1:"+ eqListHeaderColumnEnd +"1"), _service, _spreadsheet_id)
	_requests = compileFormatCells_HoursCountFormat(actualFacilityEqSheetID, _requests, gs.A1NotationToRC("C2:Z100"), _service, _spreadsheet_id)

	vals_equip = ["Equipment Name:"]
	vals_equip_type = ["Equipment Type:"]
	plannedEquipmentSheetID = str(int(_plannedFacilityEqSheetID) + 1)
	for work_center in _facility.work_center_list:

		#For each work center that will be need a time study sheet generated for it, generate <Work Center Name> Tab If Does Not Already Exist
		if(work_center.name in TIME_STUDY_WORK_CENTERS):

			[_requests, _data, _plannedFacilityWorkCenterTimeStudySheetID] = generate_time_study_sheet(work_center, _facility, _spreadsheetInfo, _requests, _data, _service, _spreadsheet_id, _plannedFacilityWorkCenterTimeStudySheetID)
		
		for equipment in work_center.workcenter_equipment_list:
			if(equipment.name not in vals_equip):
				vals_equip.append(equipment.name)
				vals_equip_type.append(equipment.type)

				#For each piece of equipment in facility, generate <Equipment Name> Tab If Does Not Already Exist and Hide
				[_requests, _data, plannedEquipmentSheetID] = generate_equipment_sheet(equipment, work_center, _facility, facility_network, _spreadsheetInfo, _requests, _data, _service, _spreadsheet_id, _admin_users, plannedEquipmentSheetID, _masterSheetColumns)

	_requests = compileConditionalFormattingBgGradient(actualFacilityEqSheetID, _requests, gs.A1NotationToRC("C2:"+ eqListHeaderColumnEnd + str(len(vals_equip))), _service, _spreadsheet_id, FACILITY_HOURS[_facility.name.capitalize()]['min'], FACILITY_HOURS[_facility.name.capitalize()]['mid'], FACILITY_HOURS[_facility.name.capitalize()]['max'])				
	_requests = gs.compileChangeColumnWidth(actualFacilityEqSheetID, _requests, gs.A1NotationToRC("A1"), _service, _spreadsheet_id, pixelWidth = 230)

	_data.append({
			"range": _facility.name.capitalize() + ' Equipment List!A1:B400',
			"values": [vals_equip,vals_equip_type],
			"majorDimension": "COLUMNS"
		})

	_data.append({
			"range": _facility.name.capitalize() + ' Equipment List!C1:'+ eqListHeaderColumnEnd +'1',
			"values": [facility_network.production_days],
			"majorDimension": "ROWS"
		})

	_requests = gs.compileRepeatCellRequest(actualFacilityEqSheetID, _requests, gs.A1NotationToRC("C2:"+ eqListHeaderColumnEnd + str(len(vals_equip))), _service, _spreadsheet_id, "=sumifs('Master Task List'!$" + _masterSheetColumns["expected_duration"] + "$6:$" + _masterSheetColumns["expected_duration"] + "$1400,\
			'Master Task List'!$" + _masterSheetColumns["equipment"][0] + "$6:$" + _masterSheetColumns["equipment"][0] + "$1400, \"=\"&$A2,'Master Task List'!$" + _masterSheetColumns["facility"] + "$6:$" + _masterSheetColumns["facility"] + "$1400, \"=" + _facility.name.capitalize() + "\",'Master Task List'!$" + _masterSheetColumns["day_allocated"] + "$6:$" + _masterSheetColumns["day_allocated"] + "$1400,\"=\"&C$1)+sumifs('Master Task List'!$" + _masterSheetColumns["expected_duration"] + "$6:$" + _masterSheetColumns["expected_duration"] + "$1400,'Master Task List'!$" + _masterSheetColumns["equipment"][1] + "$6:$" + _masterSheetColumns["equipment"][1] + "$1400, \"=\"&$A2,\
			'Master Task List'!$" + _masterSheetColumns["facility"] + "$6:$" + _masterSheetColumns["facility"] + "$1400, \"=" + _facility.name.capitalize() + "\",'Master Task List'!$" + _masterSheetColumns["day_allocated"] + "$6:$" + _masterSheetColumns["day_allocated"] + "$1400,\"=\"&C$1)+sumifs('Master Task List'!$" + _masterSheetColumns["expected_duration"] + "$6:$" + _masterSheetColumns["expected_duration"] + "$1400,'Master Task List'!$" + _masterSheetColumns["equipment"][2] + "$6:$" + _masterSheetColumns["equipment"][2] + "$1400, \"=\"&$A2,'Master Task List'!$" + _masterSheetColumns["facility"] + "$6:$" + _masterSheetColumns["facility"] + "$1400, \"=" + _facility.name.capitalize() + "\",\
			'Master Task List'!$" + _masterSheetColumns["day_allocated"] + "$6:$" + _masterSheetColumns["day_allocated"] + "$1400,\"=\"&C$1)+sumifs('Master Task List'!$" + _masterSheetColumns["expected_duration"] + "$6:$" + _masterSheetColumns["expected_duration"] + "$1400,'Master Task List'!$" + _masterSheetColumns["equipment"][3] + "$6:$" + _masterSheetColumns["equipment"][3] + "$1400, \"=\"&$A2,'Master Task List'!$" + _masterSheetColumns["facility"] + "$6:$" + _masterSheetColumns["facility"] + "$1400, \"=" + _facility.name.capitalize() + "\",'Master Task List'!$" + _masterSheetColumns["day_allocated"] + "$6:$" + _masterSheetColumns["day_allocated"] + "$1400,\"=\"&C$1)+\
			sumifs('Master Task List'!$" + _masterSheetColumns["expected_duration"] + "$6:$" + _masterSheetColumns["expected_duration"] + "$1400,'Master Task List'!$" + _masterSheetColumns["equipment"][4] + "$6:$" + _masterSheetColumns["equipment"][4] + "$1400, \"=\"&$A2,'Master Task List'!$" + _masterSheetColumns["facility"] + "$6:$" + _masterSheetColumns["facility"] + "$1400, \"=" + _facility.name.capitalize() + "\",'Master Task List'!$" + _masterSheetColumns["day_allocated"] + "$6:$" + _masterSheetColumns["day_allocated"] + "$1400,\"=\"&C$1)")

	return [_requests, _data, _plannedFacilityEqSheetID, _plannedFacilityWorkCenterTimeStudySheetID]

def generate_facility_priorities(_facility, facility_network, _spreadsheetInfo, _requests, _data, _service, _spreadsheet_id, _admin_users, _plannedFacilityPrioritiesSheetID):
	[_requests, actualFacilityPrioritiesSheetID] = gs.compileAddSheetRequest(_facility.name.capitalize() + " Priorities", gs.COLOR_PURPLE, _spreadsheetInfo, _requests, _service, _spreadsheet_id, sheetID = _plannedFacilityPrioritiesSheetID)
	_plannedFacilityPrioritiesSheetID = str(int(_plannedFacilityPrioritiesSheetID)+10000)
			
	try:
		sheet = next(sheet for sheet in _spreadsheetInfo['sheets'] if sheet['properties']['sheetId'] == actualFacilityPrioritiesSheetID)
		_requests = gs.compileProtectSheet(sheet, _requests, _service, _spreadsheet_id, gs.A1NotationToRC("B1"), _admin_users)
		tabRows = sheet['properties']['gridProperties']['rowCount']
		tabColumns = sheet['properties']['gridProperties']['columnCount']
	except StopIteration:
		_requests = gs.compileProtectSheet("", _requests, _service, _spreadsheet_id, gs.A1NotationToRC("B1"), _admin_users, newSheet = {"newSheet": True, "id": actualFacilityPrioritiesSheetID})
		tabRows = 1000
		tabColumns = 26

	equipmentCt = sum(len(work_center.workcenter_equipment_list) for work_center in _facility.work_center_list)

	if(tabColumns < equipmentCt*4+2):
		_requests = gs.compileAddDimensionToSheet(actualFacilityPrioritiesSheetID, _requests, _service, _spreadsheet_id, dimension = "COLUMNS", countToAppend = (equipmentCt*4 + 2 - tabColumns))

	_requests = gs.compileWrapStrategy(actualFacilityPrioritiesSheetID, _requests, [0, 200, 0, equipmentCt*4 + 1], _service, _spreadsheet_id, wrapStrategy = "WRAP")

	_data.append({
			"range": _facility.name.capitalize() + ' Priorities!A1:B1',
			"values": [["Day Selected", "Sunday"]],
			"majorDimension": "ROWS"
		})

	priorities_days = [day for day in facility_network.production_days]
	priorities_days.append("All")

	_requests = gs.compileDataValidationV1(actualFacilityPrioritiesSheetID, _requests, gs.A1NotationToRC("B1"), _service, _spreadsheet_id, priorities_days, dv_type = "ONE_OF_LIST", strict = True)
	_requests = gs.compileBordersAll(actualFacilityPrioritiesSheetID, _requests, gs.A1NotationToRC("A1:B1"), _service, _spreadsheet_id)
	_requests = gs.compileBatchColorCellRangeColorV1(actualFacilityPrioritiesSheetID, _requests, gs.A1NotationToRC("A1"), _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_GRAY)
	_requests = gs.compileBatchColorCellRangeColorV1(actualFacilityPrioritiesSheetID, _requests, gs.A1NotationToRC("B1"), _service, _spreadsheet_id, bgColor = gs.COLOR_YELLOW)
	
	columnStart_category = 1
	columnEnd_category = 1
	for work_center_category in _facility.work_center_category_list:
		priorities_wcs = []
		priorities_eq_list = []
		columnStart_wc = 1*columnStart_category
		columnEnd_wc = 1*columnStart_category
		if(work_center_category not in WORK_CENTER_CATEGORY_DISPLAY):
			continue
		for work_center in _facility.work_center_list:
			if(work_center.category == work_center_category):
				priorities_wcs.append({"name": work_center.name, "size": len(work_center.workcenter_equipment_list)})
				priorities_eq_list.extend([equipment.name for equipment in work_center.workcenter_equipment_list])

				columnEnd_category += 4*len(work_center.workcenter_equipment_list)

		if(columnStart_category != columnEnd_category):

			#Build Work Center Category Header
			_requests = gs.compileBatchColorCellRangeColorV1(actualFacilityPrioritiesSheetID, _requests, [2, 3, columnStart_category, columnEnd_category], _service, _spreadsheet_id, bgColor = WORK_CENTER_CATEGORY_DISPLAY[work_center_category])
			_requests = gs.compileMergeCells(actualFacilityPrioritiesSheetID, _requests, [2, 3, columnStart_category, columnEnd_category], _service, _spreadsheet_id, mergeType = "MERGE_ROWS")
			_requests = gs.compileFormatCells_TextFormat(actualFacilityPrioritiesSheetID, _requests, [2, 3, columnStart_category, columnEnd_category], _service, _spreadsheet_id, textFormat = {'bold': True, 'fontFamily': 'Arial', 'fontSize': 12})
			_requests = gs.compileFormatCells_HorizontalAlign(actualFacilityPrioritiesSheetID, _requests, [2, 3, columnStart_category, columnEnd_category], _service, _spreadsheet_id, horizAlign = "CENTER")
			_requests = gs.compileBordersBox(actualFacilityPrioritiesSheetID, _requests, [2, 5, columnStart_category, columnEnd_category], _service, _spreadsheet_id)
			_data.append({
					"range": _facility.name.capitalize() + ' Priorities!' + gs.columnIndexToLetter(columnStart_category+1) + str(3),
					"values": [[work_center_category]],
					"majorDimension": "COLUMNS"
				})

		for work_center in priorities_wcs:
			columnEnd_wc += 4*work_center["size"]
			#Build Work Center Header
			_requests = gs.compileBatchColorCellRangeColorV1(actualFacilityPrioritiesSheetID, _requests, [3, 4, columnStart_wc, columnEnd_wc], _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_GRAY)
			_requests = gs.compileMergeCells(actualFacilityPrioritiesSheetID, _requests, [3, 4, columnStart_wc, columnEnd_wc], _service, _spreadsheet_id, mergeType = "MERGE_ROWS")
			_requests = gs.compileFormatCells_TextFormat(actualFacilityPrioritiesSheetID, _requests, [3, 4, columnStart_wc, columnEnd_wc], _service, _spreadsheet_id, textFormat = {'bold': True, 'fontFamily': 'Arial', 'fontSize': 11})
			_requests = gs.compileFormatCells_HorizontalAlign(actualFacilityPrioritiesSheetID, _requests, [3, 4, columnStart_wc, columnEnd_wc], _service, _spreadsheet_id, horizAlign = "CENTER")
			_requests = gs.compileBordersBox(actualFacilityPrioritiesSheetID, _requests, [3, 4, columnStart_wc, columnEnd_wc], _service, _spreadsheet_id)
			_data.append({
					"range": _facility.name.capitalize() + ' Priorities!' + gs.columnIndexToLetter(columnStart_wc+1) + str(4),
					"values": [[work_center["name"]]],
					"majorDimension": "COLUMNS"
				})

			columnStart_wc = columnEnd_wc
	#_requests = gs.compileChangeColumnWidth(actualFacilityPrioritiesSheetID, _requests, gs.A1NotationToRC("B1:" + gs.columnIndexToLetter(columnEnd_category+1) + "1"), _service, _spreadsheet_id, pixelWidth = 180)

		counter = columnStart_category
		for equipment in priorities_eq_list:
			#for prod_day in enumerate(facility_network.production_days):
			_requests = gs.compileFormatCells_TextFormat(actualFacilityPrioritiesSheetID, _requests, [4, 5, counter, counter + 1], _service, _spreadsheet_id, textFormat = {'bold': True, 'fontFamily': 'Arial', 'fontSize': 10})
			_requests = gs.compileFormatCells_HorizontalAlign(actualFacilityPrioritiesSheetID, _requests, [4, 5, counter, counter + 1], _service, _spreadsheet_id, horizAlign = "CENTER")
			_requests = gs.compileBatchColorCellRangeColorV1(actualFacilityPrioritiesSheetID, _requests, [4, 5, counter + 1, counter + 4], _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_BLUE)
			_requests = gs.compileBordersAll(actualFacilityPrioritiesSheetID, _requests, [4, 5, counter, counter + 4], _service, _spreadsheet_id)
			_requests = gs.compileChangeColumnWidth(actualFacilityPrioritiesSheetID, _requests, [4, 5, counter, counter + 1], _service, _spreadsheet_id, pixelWidth = 180)
			_requests = gs.compileChangeColumnWidth(actualFacilityPrioritiesSheetID, _requests, [4, 5, counter + 1, counter + 2], _service, _spreadsheet_id, pixelWidth = 46)
			_requests = gs.compileChangeColumnWidth(actualFacilityPrioritiesSheetID, _requests, [4, 5, counter + 2, counter + 3], _service, _spreadsheet_id, pixelWidth = 67)
			_requests = gs.compileChangeColumnWidth(actualFacilityPrioritiesSheetID, _requests, [4, 5, counter + 3, counter + 4], _service, _spreadsheet_id, pixelWidth = 52)
			_requests = gs.compileBordersBox(actualFacilityPrioritiesSheetID, _requests, [5, 105, counter, counter + 4], _service, _spreadsheet_id)

			taskArrFormula = "="
			dataArrFormula = "="
			for prod_day in enumerate(facility_network.production_days):
				taskArrFormula += ("if($B$1=\"" + prod_day[1] + "\",arrayformula(\'" +  _facility.name.capitalize() + ": " + equipment + "\'!"+ gs.columnIndexToLetter(5*prod_day[0] + 1 + 5*len(facility_network.production_days)) + "2:"+ gs.columnIndexToLetter(5*prod_day[0] + 1 + 5*len(facility_network.production_days)) + "100),")
				dataArrFormula += ("if($B$1=\"" + prod_day[1] + "\",arrayformula(\'" +  _facility.name.capitalize() + ": " + equipment + "\'!"+ gs.columnIndexToLetter(5*prod_day[0] + 3 + 5*len(facility_network.production_days)) + "2:"+ gs.columnIndexToLetter(5*prod_day[0] + 5 + 5*len(facility_network.production_days)) + "100),")
			taskArrFormula += "if($B$1=\"All\",transpose(split(textjoin(\"_\",TRUE"
			for prod_day in enumerate(facility_network.production_days):
				taskArrFormula += (",arrayformula(\'" +  _facility.name.capitalize() + ": " + equipment + "\'!"+ gs.columnIndexToLetter(5*prod_day[0] + 1 + 5*len(facility_network.production_days)) + "2:"+ gs.columnIndexToLetter(5*prod_day[0] + 1 + 5*len(facility_network.production_days)) + "100)")
			taskArrFormula += "),\"_\")),\"\")"
			dataArrFormula += "\"\""
			for prod_day in facility_network.production_days:
				taskArrFormula += (")")
				dataArrFormula += (")")

			_data.append({
					"range": _facility.name.capitalize() + ' Priorities!' + gs.columnIndexToLetter(counter+1) + str(5) + ":" + gs.columnIndexToLetter(counter+4) + str(6),
					"values": [[equipment, taskArrFormula],
						["Man Hours", dataArrFormula],
						["People Assigned"],
						["Project Time"]],
					"majorDimension": "COLUMNS"
				})
	
			counter += 4
	
		_data.append({
				"range": _facility.name.capitalize() + ' Priorities!A' + str(5) + ":A" + str(6),
				"values": [["Priority","1"]],
				"majorDimension": "COLUMNS"
			})

		_requests = gs.compileRepeatCellRequest(actualFacilityPrioritiesSheetID, _requests, gs.A1NotationToRC("A7:A105"), _service, _spreadsheet_id, "=A6+1")

		columnEnd_category += 1
		columnStart_category = columnEnd_category


	return [_requests, _data, _plannedFacilityPrioritiesSheetID]

def generate_count_correction(_spreadsheetInfo, _requests, _data, _service, _spreadsheet_id, _sheetIDs, _planning_users, _productionRecord):
	[requests, countCorrectionSheetId] = gs.compileAddSheetRequest("Count Correction", gs.COLOR_YELLOW, _spreadsheetInfo, _requests, _service, _spreadsheet_id, sheetID = _sheetIDs['count_correction'])
		#Protect and Hide Sheet
	try:
		sheet = next(sheet for sheet in _spreadsheetInfo['sheets'] if sheet['properties']['sheetId'] == countCorrectionSheetId)
		_requests = gs.compileProtectSheet(sheet, _requests, _service, _spreadsheet_id, gs.A1NotationToRC("A1"), _planning_users)
		tabRows = sheet['properties']['gridProperties']['rowCount']
		tabColumns = sheet['properties']['gridProperties']['columnCount']
	except StopIteration:
		_requests = gs.compileProtectSheet("", _requests, _service, _spreadsheet_id, gs.A1NotationToRC("A1"), _planning_users, newSheet = {"newSheet": True, "id": countCorrectionSheetId})
		tabRows = 1000
		tabColumns = 26
	_requests = gs.compileHideSheet(countCorrectionSheetId, _requests)

	originalMealCounts = [[str(meal['mealCode']) for meal in _productionRecord['cycles']['1']['meals']]]

	for cycle in _productionRecord['cycles']:
		originalMealCounts.append([meal['totalMeals'] for meal in _productionRecord['cycles'][cycle]['meals']])

	_data.append({
			"range": 'Count Correction!A2:C83',
			"values": originalMealCounts,
			"majorDimension": "COLUMNS"
		})


	return [_requests, _data]

def generate_master_tab(facility_network, _tasks, _spreadsheetInfo, _requests, _data, _service, _spreadsheet_id, _sheetIDs, _masterSheetColumns, _master_columns_sorted, _tovala_subtype_info_list, _planning_users, _mod_plus_users, _execution_users, _productionRecord, _lenTaskCompatibilityList = 26, _maxWorkCenterListColumns = 26, _verbosity = VL_WARNING):
	[_requests, masterTaskSheetId] = gs.compileAddSheetRequest("Master Task List", gs.COLOR_RED, _spreadsheetInfo, _requests, _service, _spreadsheet_id, sheetID = _sheetIDs['master_task_list'])
		#Protect and Hide Sheet
	try:
		masterTaskSheet = next(sheet for sheet in _spreadsheetInfo['sheets'] if sheet['properties']['sheetId'] == masterTaskSheetId)
		_requests = gs.compileProtectSheet(masterTaskSheet, _requests, _service, _spreadsheet_id, gs.A1NotationToRC("A1"), _planning_users)
		tabRows = masterTaskSheet['properties']['gridProperties']['rowCount']
		tabColumns = masterTaskSheet['properties']['gridProperties']['columnCount']
	except StopIteration:
		raise IndexError("INVALID TEMPLATE: Production Planner Template Was Missing a Master Tab")
	for user_to_add in _planning_users:
		_requests = gs.compileUpdateSheetProtectedRangesAddUser(masterTaskSheet, _requests, _service, _spreadsheet_id, user_to_add)

	_data.append({
			"range": 'Master Task List!B1',
			"values": [[_productionRecord['termID']]],
			"majorDimension": "COLUMNS"
		})
	try:
		_data.append({
				"range": 'Master Task List!B2',
				"values": [["=(Date(2017,1,1)+(24*" + str(next(day[0] for day in enumerate(DAYS) if day[1] == facility_network.production_days[0])) + "+(B1-1)*168)*3600)/86400+date(2017,1,1)"]],
				"majorDimension": "COLUMNS"
			})
		_data.append({
				"range": 'Master Task List!B3',
				"values": [["=(Date(2017,1,1)+(24*" + str(next(day[0] for day in enumerate(DAYS) if day[1] == facility_network.production_days[0]) + len(facility_network.production_days) - 1) + "+(B1-1)*168)*3600)/86400+date(2017,1,1)"]],
				"majorDimension": "COLUMNS"
			})
	except StopIteration:
		raise IndexError("INVALID PRODUCTION DAYS: First production day not found in DAYS")

	lastColumn = _master_columns_sorted[-1]['column']

	_data.append({
			"range": 'Master Task List!A5:' + lastColumn + "5",
			"values": [[column['title'] for column in _master_columns_sorted]],
			"majorDimension": "ROWS"
		})

	_requests = gs.compileMergeCells(masterTaskSheetId, _requests, gs.A1NotationToRC("AQ1:AS3"), _service, _spreadsheet_id, mergeType = "MERGE_ROWS")
	_requests = gs.compileBatchColorCellRangeColorV1(masterTaskSheetId, _requests, gs.A1NotationToRC("AQ1:AS3"), _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_GRAY)
	_requests = gs.compileBatchColorCellRangeColorV1(masterTaskSheetId, _requests, gs.A1NotationToRC("AT1:AT3"), _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_BLUE)
	_requests = gs.compileBordersAll(masterTaskSheetId, _requests, gs.A1NotationToRC("AQ1:AT3"), _service, _spreadsheet_id)

	c1Count = 0
	c2Count = 0
	for cycle in _productionRecord['cycles']:
		if cycle == "1":
			for meal in _productionRecord['cycles'][cycle]['meals']:
				c1Count += meal['totalMeals']
		if cycle == "2":
			for meal in _productionRecord['cycles'][cycle]['meals']:
				c2Count += meal['totalMeals']

	_data.append({
			"range": 'Master Task List!AQ1:AQ3',
			"values": [["Term Meal Count", "Cycle 1 Meal Count", "Cycle 2 Meal Count"]],
			"majorDimension": "COLUMNS"
		})

	_data.append({
			"range": 'Master Task List!AT1:AT3',
			"values": [[str(c1Count+c2Count),str(c1Count),str(c2Count)]],
			"majorDimension": "COLUMNS"
		})

	for i in range(9):
		equipmentPossibleColumn = next(column['column'] for column in _master_columns_sorted if column['title'] == ("Equipment Possible_" + str(i)))
		_requests = gs.compileRepeatCellRequest(masterTaskSheetId, _requests, gs.A1NotationToRC(equipmentPossibleColumn + "6:" + equipmentPossibleColumn + "1000"), _service, _spreadsheet_id, "=index(indirect(G6&\" Work Centers!$B$2:$" + gs.columnIndexToLetter(_maxWorkCenterListColumns) + "$10\")," + str(i+1) + ",match(L6,indirect(G6&\" Work Centers!B1:" + gs.columnIndexToLetter(_maxWorkCenterListColumns) + "1\"),0))")
	_requests = gs.compileHideColumn(masterTaskSheetId, _requests, gs.A1NotationToRC(next(column['column'] for column in _master_columns_sorted if column['title'] == ("Equipment Possible_0")) + "1:" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Equipment Possible_8")) + "1"), _service, _spreadsheet_id, hide = True)

	#_requests = gs.compileDataValidationV1(masterTaskSheetId, _requests, gs.A1NotationToRC(_masterSheetColumns['equipment'][0] + "6:" + _masterSheetColumns['equipment'][4] + "1400"), _service, _spreadsheet_id, ["='Master Task List'!$" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Equipment Possible_0")) + "6:$" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Equipment Possible_8")) + "6"], dv_type = "ONE_OF_RANGE", strict = False)
	#^^ this request can't be completed through the API and must be completed through a template instead

	for i in range(4):
		workCenterPossibleColumn = next(column['column'] for column in _master_columns_sorted if column['title'] == ("Work Centers Possible_" + str(i)))
		_requests = gs.compileRepeatCellRequest(masterTaskSheetId, _requests, gs.A1NotationToRC(workCenterPossibleColumn + "6:" + workCenterPossibleColumn + "1000"), _service, _spreadsheet_id, "=index(if(G6=\"" + facility_network.facilityList[0].name.capitalize() + "\",\'" + facility_network.name.capitalize() + " Task Compatibility\'!$B$2:$" + gs.columnIndexToLetter(_lenTaskCompatibilityList) + "$5,\'" + facility_network.name.capitalize() + " Task Compatibility\'!$B$6:$" + gs.columnIndexToLetter(_lenTaskCompatibilityList) + "$9)," + str(i+1) + ",match(E6,\'" + facility_network.name.capitalize() + " Task Compatibility\'!$B$1:$" + gs.columnIndexToLetter(_lenTaskCompatibilityList) + "$1,0))")
	_requests = gs.compileHideColumn(masterTaskSheetId, _requests, gs.A1NotationToRC(next(column['column'] for column in _master_columns_sorted if column['title'] == ("Work Centers Possible_0")) + "1:" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Work Centers Possible_3")) + "1"), _service, _spreadsheet_id, hide = True)

	#_requests = gs.compileDataValidationV1(masterTaskSheetId, _requests, gs.A1NotationToRC(next(column['column'] for column in _master_columns_sorted if column['title'] == ("Work Center")) + "6:" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Work Center")) + "1400"), _service, _spreadsheet_id, ["='Master Task List'!$" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Work Centers Possible_0")) + "6:$" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Work Centers Possible_3")) + "6"], dv_type = "ONE_OF_RANGE", strict = False)
	#^^ this request can't be completed through the API and must be completed through a template instead

	_requests = gs.compileDataValidationV1(masterTaskSheetId, _requests, gs.A1NotationToRC(_masterSheetColumns['facility'] + "6:" + _masterSheetColumns['facility'] + "1400"), _service, _spreadsheet_id, [facility.name.capitalize() for facility in facility_network.facilityList], dv_type = "ONE_OF_LIST", strict = False)
	_requests = gs.compileRepeatCellRequest(masterTaskSheetId, _requests, gs.A1NotationToRC(next(column['column'] for column in _master_columns_sorted if column['title'] == ("Equipment Type_0")) + "6:" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Equipment Type_4")) + "1400"), _service, _spreadsheet_id, "=ifna(vlookup(" + _masterSheetColumns['equipment'][0] + "6,indirect($" + _masterSheetColumns['facility'] + "6&\" Equipment List!A2:70\"),2,FALSE),\"\")")

	_requests = gs.compileMergeCells(masterTaskSheetId, _requests, gs.A1NotationToRC(_masterSheetColumns['equipment'][0] + "5:" + _masterSheetColumns['equipment'][4] + "5"), _service, _spreadsheet_id, mergeType = "MERGE_ROWS")
	_requests = gs.compileBordersBox(masterTaskSheetId, _requests, gs.A1NotationToRC("A5:" + lastColumn + "5"), _service, _spreadsheet_id)
	_requests = gs.compileBatchColorCellRangeColorV1(masterTaskSheetId, _requests, gs.A1NotationToRC("A5:" + lastColumn + "5"), _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_GRAY)
	_requests = gs.compileBatchColorCellRangeColorV1(masterTaskSheetId, _requests, gs.A1NotationToRC("A1:A3"), _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_GRAY)
	_requests = gs.compileBatchColorCellRangeColorV1(masterTaskSheetId, _requests, gs.A1NotationToRC("B1:B3"), _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_BLUE)
	_requests = gs.compileBordersBox(masterTaskSheetId, _requests, gs.A1NotationToRC("A1:A3"), _service, _spreadsheet_id)
	_requests = gs.compileBordersBox(masterTaskSheetId, _requests, gs.A1NotationToRC("B1:B3"), _service, _spreadsheet_id)
	_requests = gs.compileBatchColorCellRangeColorV1(masterTaskSheetId, _requests, gs.A1NotationToRC(_masterSheetColumns['equipment'][0] + "5:" + _masterSheetColumns['equipment'][4] + "5"), _service, _spreadsheet_id, bgColor = gs.COLOR_DARK_GRAY_2)
	_requests = gs.compileFormatCells_TextFormat(masterTaskSheetId, _requests, gs.A1NotationToRC("A5:" + lastColumn + "5"), _service, _spreadsheet_id, textFormat = {'bold': True, 'fontFamily': 'Calibri', 'fontSize': 11})
	_requests = gs.compileHideColumn(masterTaskSheetId, _requests, gs.A1NotationToRC(next(column['column'] for column in _master_columns_sorted if column['title'] == ("Equipment Type_0")) + "1:" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Equipment Type_4")) + "1"), _service, _spreadsheet_id, hide = True)
	_requests = gs.compileWrapStrategy(masterTaskSheetId, _requests, gs.A1NotationToRC(_masterSheetColumns['work_order'] + "6:" + _masterSheetColumns['work_order'] + "1400"), _service, _spreadsheet_id, wrapStrategy = "WRAP")
	_requests = gs.compileWrapStrategy(masterTaskSheetId, _requests, gs.A1NotationToRC("B6:B1400"), _service, _spreadsheet_id, wrapStrategy = "CLIP")

	_requests = gs.compileBordersBox(masterTaskSheetId, _requests, gs.A1NotationToRC("A6:E1400"), _service, _spreadsheet_id)
	_requests = gs.compileBordersBox(masterTaskSheetId, _requests, gs.A1NotationToRC("G6:L1400"), _service, _spreadsheet_id)
	_requests = gs.compileBatchColorCellRangeColorV1(masterTaskSheetId, _requests, gs.A1NotationToRC("G6:L1400"), _service, _spreadsheet_id, bgColor = gs.COLOR_YELLOW)
	_requests = gs.compileBordersBox(masterTaskSheetId, _requests, gs.A1NotationToRC("M6:N1400"), _service, _spreadsheet_id)

	equipment_selected_columns = gs.A1NotationToRC("X6:AB1400")
	_requests = gs.compileBordersBox(masterTaskSheetId, _requests, equipment_selected_columns, _service, _spreadsheet_id)
	_requests = gs.compileBatchColorCellRangeColorV1(masterTaskSheetId, _requests, equipment_selected_columns, _service, _spreadsheet_id, bgColor = gs.COLOR_YELLOW)
	_requests = gs.compileUpdateProtectedRangeAddUnprotectedRangeToSheet(masterTaskSheet, _requests, _service, _spreadsheet_id, equipment_selected_columns)
	_requests = gs.compileUpdateProtectedRangeAddProtectedRangeToSheet(masterTaskSheet, _requests, _service, _spreadsheet_id, equipment_selected_columns, _mod_plus_users)
	

	_requests = gs.compileBordersBox(masterTaskSheetId, _requests, gs.A1NotationToRC("AH6:" + lastColumn + "1400"), _service, _spreadsheet_id)

	pplAssignedColumn = gs.A1NotationToRC(next(column['column'] for column in _master_columns_sorted if column['title'] == ("# of People Assigned")) + "6:" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("# of People Assigned")) + "1400")
	_requests = gs.compileBatchColorCellRangeColorV1(masterTaskSheetId, _requests, pplAssignedColumn, _service, _spreadsheet_id, bgColor = gs.COLOR_YELLOW)
	_requests = gs.compileBordersBox(masterTaskSheetId, _requests, pplAssignedColumn, _service, _spreadsheet_id)

	estHoursColumn = gs.A1NotationToRC(next(column['column'] for column in _master_columns_sorted if column['title'] == ("Estimated Hours")) + "6:" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Estimated Hours")) + "1400")
	_requests = gs.compileBatchColorCellRangeColorV1(masterTaskSheetId, _requests, estHoursColumn, _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_BLUE)

	dayAllocatedColumn = gs.A1NotationToRC(next(column['column'] for column in _master_columns_sorted if column['title'] == ("Production Day")) + "6:" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Production Day")) + "1400")
	_requests = gs.compileBatchColorCellRangeColorV1(masterTaskSheetId, _requests, dayAllocatedColumn, _service, _spreadsheet_id, bgColor = gs.COLOR_YELLOW)
	_requests = gs.compileBordersBox(masterTaskSheetId, _requests, dayAllocatedColumn, _service, _spreadsheet_id)
	_requests = gs.compileDataValidationV1(masterTaskSheetId, _requests, dayAllocatedColumn, _service, _spreadsheet_id, facility_network.production_days, dv_type = "ONE_OF_LIST", strict = False)
	_requests = gs.compileUpdateProtectedRangeAddUnprotectedRangeToSheet(masterTaskSheet, _requests, _service, _spreadsheet_id, dayAllocatedColumn)
	_requests = gs.compileUpdateProtectedRangeAddProtectedRangeToSheet(masterTaskSheet, _requests, _service, _spreadsheet_id, dayAllocatedColumn, _mod_plus_users)

	priorityColumn = gs.A1NotationToRC(next(column['column'] for column in _master_columns_sorted if column['title'] == ("Priority")) + "6:" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Priority")) + "1400")
	_requests = gs.compileBatchColorCellRangeColorV1(masterTaskSheetId, _requests, priorityColumn, _service, _spreadsheet_id, bgColor = gs.COLOR_YELLOW)
	_requests = gs.compileDataValidationV1(masterTaskSheetId, _requests, priorityColumn, _service, _spreadsheet_id, ["1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20"], dv_type = "ONE_OF_LIST", strict = False)
	_requests = gs.compileBordersBox(masterTaskSheetId, _requests, priorityColumn, _service, _spreadsheet_id)
	_requests = gs.compileUpdateProtectedRangeAddUnprotectedRangeToSheet(masterTaskSheet, _requests, _service, _spreadsheet_id, priorityColumn)
	_requests = gs.compileUpdateProtectedRangeAddProtectedRangeToSheet(masterTaskSheet, _requests, _service, _spreadsheet_id, priorityColumn, _mod_plus_users)


	_requests = gs.compileHideColumn(masterTaskSheetId, _requests, gs.A1NotationToRC("F1"), _service, _spreadsheet_id, hide = True)
	_requests = gs.compileChangeColumnWidth(masterTaskSheetId, _requests, gs.A1NotationToRC("D1"), _service, _spreadsheet_id, pixelWidth = 300)
	_requests = gs.compileFormatCells_NumberFormat(masterTaskSheetId, _requests, gs.A1NotationToRC("AV6:AV1400"), _service, _spreadsheet_id, numFormat = "DATE")

	statusColumn = gs.A1NotationToRC(next(column['column'] for column in _master_columns_sorted if column['title'] == ("Status")) + "6:" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Status")) + "1400")
	_requests = gs.compileDataValidationV1(masterTaskSheetId, _requests, statusColumn, _service, _spreadsheet_id, ["In Progress", "Complete", "Short"], dv_type = "ONE_OF_LIST", strict = False)

	prodColumns = gs.A1NotationToRC(next(column['column'] for column in _master_columns_sorted if column['title'] == ("Status")) + "6:" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Produced Quantity")) + "1400")
	_requests = gs.compileBatchColorCellRangeColorV1(masterTaskSheetId, _requests, prodColumns, _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_PURPLE_3)
	_requests = gs.compileUpdateProtectedRangeAddUnprotectedRangeToSheet(masterTaskSheet, _requests, _service, _spreadsheet_id, prodColumns)
	_requests = gs.compileUpdateProtectedRangeAddProtectedRangeToSheet(masterTaskSheet, _requests, _service, _spreadsheet_id, prodColumns, _execution_users)

	initialsColumn = gs.A1NotationToRC(next(column['column'] for column in _master_columns_sorted if column['title'] == ("Initials")) + "6:" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Initials")) + "1400")
	_requests = gs.compileBatchColorCellRangeColorV1(masterTaskSheetId, _requests, initialsColumn, _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_PURPLE_3)
	initial_names = [user['name'] for user in facility_network.staffing['mods']]
	initial_names.extend([user['name'] for user in facility_network.staffing['supervisors']])
	initial_names.extend([user['name'] for user in facility_network.staffing['plant_managers']])
	_requests = gs.compileDataValidationV1(masterTaskSheetId, _requests, initialsColumn, _service, _spreadsheet_id, initial_names, dv_type = "ONE_OF_LIST", strict = False)
	_requests = gs.compileBordersBox(masterTaskSheetId, _requests, initialsColumn, _service, _spreadsheet_id)
	_requests = gs.compileUpdateProtectedRangeAddUnprotectedRangeToSheet(masterTaskSheet, _requests, _service, _spreadsheet_id, initialsColumn)
	_requests = gs.compileUpdateProtectedRangeAddProtectedRangeToSheet(masterTaskSheet, _requests, _service, _spreadsheet_id, initialsColumn, _execution_users)

	_requests = gs.compileConditionalFormattingBgColor(masterTaskSheetId, _requests, gs.A1NotationToRC("A6:AZ1400"), _service, _spreadsheet_id, "=$" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Status")) + "6=\"Complete\"", gs.COLOR_LIGHT_GREEN)
	_requests = gs.compileConditionalFormattingBgColor(masterTaskSheetId, _requests, gs.A1NotationToRC("A6:AZ1400"), _service, _spreadsheet_id, "=$" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Status")) + "6=\"In Progress\"", gs.COLOR_LIGHT_YELLOW_3)
	_requests = gs.compileConditionalFormattingBgColor(masterTaskSheetId, _requests, gs.A1NotationToRC("A6:AZ1400"), _service, _spreadsheet_id, "=$" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Status")) + "6=\"Short\"", gs.COLOR_LIGHT_RED)
	
	time_start_column_letters = next(column['column'] for column in _master_columns_sorted if column['title'] == ("Start Time\n(HR:MM AM/PM)"))
	time_breaks_column_letters = next(column['column'] for column in _master_columns_sorted if column['title'] == ("Breaks/Lunch\n(Minutes)"))
	time_end_column_letters = next(column['column'] for column in _master_columns_sorted if column['title'] == ("End Time"))
	team_column_letters = next(column['column'] for column in _master_columns_sorted if column['title'] == ("Team Members"))
	_requests = gs.compileBatchColorCellRangeColorV1(masterTaskSheetId, _requests, gs.A1NotationToRC(time_start_column_letters + "6:" + team_column_letters + "1400"), _service, _spreadsheet_id, bgColor = gs.COLOR_YELLOW)
	_requests = gs.compileBatchColorCellRangeColorV1(masterTaskSheetId, _requests, gs.A1NotationToRC(time_end_column_letters + "6:" + time_end_column_letters + "1400"), _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_BLUE)
	_requests = gs.compileBordersBox(masterTaskSheetId, _requests, gs.A1NotationToRC(time_end_column_letters + "6:" + time_end_column_letters + "1400"), _service, _spreadsheet_id)
	_requests = gs.compileRepeatCellRequest(masterTaskSheetId, _requests, gs.A1NotationToRC(time_end_column_letters + "6:" + time_end_column_letters + "1400"), _service, _spreadsheet_id, "=if(isblank(" + time_start_column_letters + "6),\"\"," + time_start_column_letters + "6+" + time_breaks_column_letters + "6/(24*60)+" + _masterSheetColumns['estHours'] + "6/24)")

	_requests = gs.compileWrapStrategy(masterTaskSheetId, _requests, gs.A1NotationToRC(team_column_letters + "6:" + team_column_letters + "1400"), _service, _spreadsheet_id, wrapStrategy = "WRAP")
	_requests = gs.compileChangeColumnWidth(masterTaskSheetId, _requests, gs.A1NotationToRC(team_column_letters + "6:" + team_column_letters + "1400"), _service, _spreadsheet_id, pixelWidth = 230)

	_requests = gs.compileRepeatCellRequest(masterTaskSheetId, _requests, gs.A1NotationToRC(_masterSheetColumns['composite'] + "6:" + _masterSheetColumns['composite'] + "1400"), _service, _spreadsheet_id, "=textjoin(\"-\", FALSE, " + _masterSheetColumns["work_order"] + "6, " + _masterSheetColumns["day_allocated"] + "6, " + _masterSheetColumns["meals"] + "6,\"Cycle\"," + _masterSheetColumns["cycle"] + "6)")


	short_surplus_column_letters = next(column['column'] for column in _master_columns_sorted if column['title'] == ("Short/Surplus"))
	_requests = gs.compileBatchColorCellRangeColorV1(masterTaskSheetId, _requests, gs.A1NotationToRC(short_surplus_column_letters + "6:" + short_surplus_column_letters + "1400"), _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_GREEN)
	_requests = gs.compileConditionalFormattingBgColor(masterTaskSheetId, _requests, gs.A1NotationToRC(short_surplus_column_letters + "6:" + short_surplus_column_letters + "1400"), _service, _spreadsheet_id, "=$" + short_surplus_column_letters + "6<0", gs.COLOR_LIGHT_RED)
	_requests = gs.compileRepeatCellRequest(masterTaskSheetId, _requests, gs.A1NotationToRC(short_surplus_column_letters + "6:" + short_surplus_column_letters + "1400"), _service, _spreadsheet_id, "=AY6-M6")

	tasks_0 = []
	tasks_1 = []
	tasks_2 = []
	tasks_3 = []



	counter = 6
	date_allocated_string = "{\""
	for day in facility_network.production_days:
		date_allocated_string += day + "\",\""
	date_allocated_string = date_allocated_string[:-3]
	date_allocated_string += "\"},0)"

	duration = [next(column['column'] for column in _master_columns_sorted if column['title'] == ("Time for Setup")),next(column['column'] for column in _master_columns_sorted if column['title'] == ("Time for Execution")),next(column['column'] for column in _master_columns_sorted if column['title'] == ("Time for Cleanup"))]
	
	for task in _tasks:
		
		try:
			task_work_center = next(work_center for work_center in next(facility for facility in facility_network.facilityList if facility.name == task.planned_location).work_center_list if task.type in work_center.compatible_process_list)
		except StopIteration:
			task_work_center = eq.WorkCenter("",[],[],"")

		task_available_equipment = task_work_center.workcenter_equipment_list

		if(verbosity <= VL_INFORMATIONAL):
			print(task.name)
			print(task_work_center.name)
			print([equipment.name for equipment in task_available_equipment])
			if(task.type == "Kettle"):
				print(task_available_equipment[0].additional_properties)

		try:
			process_index = next(process[0] for process in enumerate([process['name'] for process in _tovala_subtype_info_list]) if process[1] == task.subtype)
		except StopIteration:
			process_index = 0
			if(_verbosity <= VL_DEBUG):
				print("WARNING: %s Not In Process Subtype Info List" % task.subtype)

		task = scheduler.equipment_scheduler(task, task_available_equipment, facility_network)[0]

		tasks_2.append(task.assigned_equipment)

		tasks_3.append([re.sub(r'<ROW>', str(counter),_tovala_subtype_info_list[process_index]['setup_workers']), re.sub(r'<ROW>', str(counter), _tovala_subtype_info_list[process_index]['assemblers']), re.sub(r'<ROW>', str(counter),  _tovala_subtype_info_list[process_index]['cooks']),\
			re.sub(r'<ROW>', str(counter), _tovala_subtype_info_list[process_index]['cleanup_workers']), re.sub(r'<ROW>', str(counter),  _tovala_subtype_info_list[process_index]['batch_quantity']), re.sub(r'<ROW>', str(counter),  _tovala_subtype_info_list[process_index]['quantity_per_batch']),\
			re.sub(r'<ROW>', str(counter), _tovala_subtype_info_list[process_index]['setup_duration']), re.sub(r'<ROW>', str(counter), _tovala_subtype_info_list[process_index]['execution_duration']), re.sub(r'<ROW>', str(counter), _tovala_subtype_info_list[process_index]['cleanup_duration']),\
			 "=" + duration[0] + str(counter) + "+" + duration[1] + str(counter) + "+" + duration[2] + str(counter),"=(AI" + str(counter) + "+AJ" + str(counter) + ")*AO" + str(counter) + "+AH" + str(counter) + "*AN" + str(counter) + "+AK" + str(counter) + "*AP" + str(counter),\
			 "=AI"+ str(counter) + "+AJ" + str(counter), "=AR"+ str(counter) + "/AS" + str(counter), task.planned_day,\
			  "=(Date(2017,1,1)+(24*match(" + _masterSheetColumns['day_allocated'] + str(counter) + "," + date_allocated_string + "-24+($B$1-1)*168)*3600)/86400+date(2017,1,1)",\
			  "","","",""])


		task_meal = "_"
		task_parents = "_"
		task_meal = task_meal.join([str(meal) for meal in task.meals])
		task_parents = task_parents.join([str(parent) for parent in task.parents])
		tasks_0.append([task.cycle, task_meal, task_parents, task.name, task.type, task.subtype, task.planned_location])
		if(task.type in OUTPUT_TASK_TYPES):
			count = np.sum([output.count for output in task.outputs])
			uom = task.outputs[0].unit
		else:
			count = np.sum([inpt.count for inpt in task.inputs])
			uom = task.inputs[0].unit
		tasks_1.append([task_work_center.name,count,uom])
		counter += 1

	_data.append({
			"range": "Master Task List!A6:G1400",
			"values": tasks_0,
			"majorDimension": "ROWS"
		})

	_data.append({
			"range": "Master Task List!L6:N1400",
			"values": tasks_1,
			"majorDimension": "ROWS"
		})

	_data.append({
			"range": "Master Task List!X6:AB1400",
			"values": tasks_2,
			"majorDimension": "ROWS"
		})

	_data.append({
			"range": "Master Task List!AH6:" + lastColumn + "1400",
			"values": tasks_3,
			"majorDimension": "ROWS"
		})

	return [_requests, _data]

def generate_meal_production_tracking_list(facility_network, _tasks, _spreadsheetInfo, _requests, _data, _service, _spreadsheet_id, _sheetIDs, _master_columns_sorted, _admin_users, _planning_users, _execution_users, _productionRecord):
	[_requests, MPTSheetId] = gs.compileAddSheetRequest("Meal Production Tracking", gs.COLOR_BLUE, _spreadsheetInfo, _requests, _service, _spreadsheet_id, sheetID = _sheetIDs['meal_production_tracking'])
		#Protect Sheet
	try:
		sheet = next(sheet for sheet in _spreadsheetInfo['sheets'] if sheet['properties']['sheetId'] == MPTSheetId)
		_requests = gs.compileProtectSheet(sheet, _requests, _service, _spreadsheet_id, gs.A1NotationToRC("A1"), _execution_users)
		_requests = gs.compileUpdateProtectedRangeAddProtectedRangeToSheet(sheet, _requests, _service, _spreadsheet_id, gs.A1NotationToRC("M1:P900"), _planning_users)
		tabRows = sheet['properties']['gridProperties']['rowCount']
		tabColumns = sheet['properties']['gridProperties']['columnCount']
	except StopIteration:
		_requests = gs.compileProtectSheet("", _requests, _service, _spreadsheet_id, gs.A1NotationToRC("A1"), _execution_users, newSheet = {"newSheet": True, "id": MPTSheetId})
		_requests = gs.compileUpdateProtectedRangeAddProtectedRangeToSheet("", _requests, _service, _spreadsheet_id, gs.A1NotationToRC("M1:P900"), _planning_users, newSheet = {"newSheet": True, "id": MPTSheetId})
		_requests = gs.compileUpdateProtectedRangeAddProtectedRangeToSheet("", _requests, _service, _spreadsheet_id, gs.A1NotationToRC("G1:H900"), _planning_users, newSheet = {"newSheet": True, "id": MPTSheetId})
		tabRows = 1000
		tabColumns = 26

	if(tabColumns < 28):
		_requests = gs.compileAddDimensionToSheet(MPTSheetId, _requests, _service, _spreadsheet_id, dimension = "COLUMNS", countToAppend = 7)

	_requests = gs.compileBordersAll(MPTSheetId, _requests, gs.A1NotationToRC("A2:AB2"), _service, _spreadsheet_id)

	mealIDs = []
	for cycle in _productionRecord['cycles']:
		for meal in _productionRecord['cycles'][cycle]['meals']:
			if(meal['mealCode'] not in [_meal['id'] for _meal in mealIDs]):
				mealIDs.append({"id": meal['mealCode'], "name": meal['shortTitle']})

	dataArray = []
	counter = 3
	task_meal = "/"
	for meal in mealIDs:

		mealStartCounter = counter

		taskCount = ["",""]
		taskFacility = ["",""]
		taskStatus = ["", ""]
		gtTaskStatus = ["", ""]
		try:
			sleevingTask_Cycle1 = next(task[0] for task in enumerate(_tasks) if task_meal.join([str(meal) for meal in task[1].meals]) == str(meal['id']) and task[1].type == "Sleeving" and str(task[1].cycle) == "1")
			garnishTrayTask_Cycle1 = next(task[0] for task in enumerate(_tasks) if task_meal.join([str(meal) for meal in task[1].meals]) == str(meal['id']) and task[1].type == "Garnish Tray Building" and str(task[1].cycle) == "1")
			taskFacility[0] = "=\'Master Task List\'!" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Facility")) + str(sleevingTask_Cycle1+6)
			taskStatus[0] = "=\'Master Task List\'!" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Status")) + str(sleevingTask_Cycle1+6)
			gtTaskStatus[0] = "=\'Master Task List\'!" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Status")) + str(garnishTrayTask_Cycle1+6)
			taskCount[0] = "=\'Master Task List\'!" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Count")) + str(sleevingTask_Cycle1+6)
		except StopIteration:
			if(verbosity<VL_INFORMATIONAL):
				print("Meal %s Does Not Have a C1 Sleeving or Garnish Tray Task" % str(meal['id']))

		try:
			sleevingTask_Cycle2 = next(task[0] for task in enumerate(_tasks) if task_meal.join([str(meal) for meal in task[1].meals]) == str(meal['id']) and task[1].type == "Sleeving" and str(task[1].cycle) == "2")
			garnishTrayTask_Cycle2 = next(task[0] for task in enumerate(_tasks) if task_meal.join([str(meal) for meal in task[1].meals]) == str(meal['id']) and task[1].type == "Garnish Tray Building" and str(task[1].cycle) == "2")
			taskFacility[1] = "=\'Master Task List\'!" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Facility")) + str(sleevingTask_Cycle2+6)
			taskStatus[1] = "=\'Master Task List\'!" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Status")) + str(sleevingTask_Cycle2+6)
			gtTaskStatus[1] = "=\'Master Task List\'!" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Status")) + str(garnishTrayTask_Cycle2+6)
			taskCount[1] = "=\'Master Task List\'!" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Count")) + str(sleevingTask_Cycle2+6)
		except StopIteration:
			if(verbosity<VL_INFORMATIONAL):
				print("Meal %s Does Not Have a C2 Sleeving or Garnish Tray Task" % str(meal['id']))

		dataArray.append(["", meal['id'], taskCount[0], taskCount[1], meal['name'],"",taskFacility[0],taskFacility[1],"Sleeve","","","",taskStatus[0],taskStatus[1],gtTaskStatus[0],gtTaskStatus[1],"=floor(C" + str(counter) + "/160)","=floor((C" + str(counter) + "-Q" + str(counter) + "*160)/16)","=mod(C" + str(counter) + ",16)","=floor(C" + str(counter) + "/640)","=floor((C" + str(counter) + "-T" + str(counter) + "*640)/64)","=mod(C" + str(counter) + ",64)","=floor(D" + str(counter) + "/160)","=floor((D" + str(counter) + "-W" + str(counter) + "*160)/16)","=mod(D" + str(counter) + ",16)","=floor(D" + str(counter) + "/640)","=floor((D" + str(counter) + "-Z" + str(counter) + "*640)/64)","=mod(D" + str(counter) + ",64)"])
		counter += 1

		portioningTasks_Cycle1 = [task[0] for task in enumerate(_tasks) if task_meal.join([str(meal) for meal in task[1].meals]) == str(meal['id']) and (task[1].type == "Tray Portioning and Sealing"\
			or task[1].type == "Clamshell Portioning" or task[1].type == "Liquid Sachet Depositing" or task[1].type == "Band Sealing" or task[1].type == "Dry Sachet Depositing" or task[1].type == "Cup Portioning") and str(task[1].cycle) == "1"]

		portioningTasks_Cycle2 = [task[0] for task in enumerate(_tasks) if task_meal.join([str(meal) for meal in task[1].meals]) == str(meal['id']) and (task[1].type == "Tray Portioning and Sealing"\
			or task[1].type == "Clamshell Portioning" or task[1].type == "Liquid Sachet Depositing" or task[1].type == "Band Sealing" or task[1].type == "Dry Sachet Depositing" or task[1].type == "Cup Portioning") and str(task[1].cycle) == "2"]

		portioningTaskComponent1 = []
		try:
			portioningTaskComponent1.append(next(task[0] for task in enumerate(_tasks) if task[0] in portioningTasks_Cycle1 and "Component 1" in task[1].name))
			portioningTasks_Cycle1.pop(portioningTasks_Cycle1.index(portioningTaskComponent1[0]))
		except StopIteration:
			portioningTaskComponent1.append(-1)
			if(verbosity <= VL_INFORMATIONAL):
				print("Meal %s Does Not Have a Component 1 in Cycle 1" % str(meal))
		if(portioningTaskComponent1[0] >= 0):
			try:
				portioningTaskComponent1.append(next(task[0] for task in enumerate(_tasks) if task[0] in portioningTasks_Cycle2 and _tasks[portioningTaskComponent1[0]].name == task[1].name))
				portioningTasks_Cycle2.pop(portioningTasks_Cycle2.index(portioningTaskComponent1[1]))
			except StopIteration:
				portioningTaskComponent1.append(-1)
				if(verbosity <= VL_INFORMATIONAL):
					print("Meal %s Does Not Have a Component 1 in Cycle 2" % str(meal))
		else:
			try:
				portioningTaskComponent1.append(next(task[0] for task in enumerate(_tasks) if task[0] in portioningTasks_Cycle2 and "Component 1" in task[1].name))
				portioningTasks_Cycle2.pop(portioningTasks_Cycle2.index(portioningTaskComponent1[1]))
			except StopIteration:
				portioningTaskComponent1.append(-1)
				if(verbosity <= VL_INFORMATIONAL):
					print("Meal %s Does Not Have a Component 1 in Cycle 2" % str(meal))

		taskName = ""
		taskPortionWeight = ""
		taskContainer = ""
		taskFacility = ["",""]
		taskStatus = ["", ""]

		for i in range(len(portioningTaskComponent1)):
			if(portioningTaskComponent1[i]>=0):
				if(taskPortionWeight == ""):
					for inpt in _tasks[portioningTaskComponent1[i]].inputs:
						if(inpt.unit == "Grams"):
							taskPortionWeight += str(round(inpt.count,2)) + " g\t"
						else:
							taskPortionWeight += str(round(inpt.count,2)) + " " + inpt.unit + "\t"
				taskFacility[i] = "=\'Master Task List\'!" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Facility")) + str(portioningTaskComponent1[i]+6)
				taskStatus[i] = "=\'Master Task List\'!" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Status")) + str(portioningTaskComponent1[i]+6)
				taskContainer = _tasks[portioningTaskComponent1[i]].outputs[0].unit
				taskName = _tasks[portioningTaskComponent1[i]].name

		dataArray.append(["", meal['id'], "", "", taskName, taskPortionWeight, taskFacility[0], taskFacility[1], taskContainer,"","","",taskStatus[0],taskStatus[1]])
		counter += 1

		portioningTaskComponent2 = []
		try:
			portioningTaskComponent2.append(next(task[0] for task in enumerate(_tasks) if task[0] in portioningTasks_Cycle1 and "Component 2" in task[1].name))
			portioningTasks_Cycle1.pop(portioningTasks_Cycle1.index(portioningTaskComponent2[0]))
		except StopIteration:
			portioningTaskComponent2.append(-1)
			if(verbosity <= VL_INFORMATIONAL):
				print("Meal %s Does Not Have a Component 2 in Cycle 1" % str(meal))
		if(portioningTaskComponent2[0] >= 0):
			try:
				portioningTaskComponent2.append(next(task[0] for task in enumerate(_tasks) if task[0] in portioningTasks_Cycle2 and _tasks[portioningTaskComponent2[0]].name == task[1].name))
				portioningTasks_Cycle2.pop(portioningTasks_Cycle2.index(portioningTaskComponent2[1]))
			except StopIteration:
				portioningTaskComponent2.append(-1)
				if(verbosity <= VL_INFORMATIONAL):
					print("Meal %s Does Not Have a Component 2 in Cycle 2" % str(meal))
		else:
			try:
				portioningTaskComponent2.append(next(task[0] for task in enumerate(_tasks) if task[0] in portioningTasks_Cycle2 and "Component 2" in task[1].name))
				portioningTasks_Cycle2.pop(portioningTasks_Cycle2.index(portioningTaskComponent2[1]))
			except StopIteration:
				portioningTaskComponent2.append(-1)
				if(verbosity <= VL_INFORMATIONAL):
					print("Meal %s Does Not Have a Component 2 in Cycle 2" % str(meal))

		taskName = ""
		taskPortionWeight = ""
		taskContainer = ""
		taskFacility = ["",""]
		taskStatus = ["", ""]

		for i in range(len(portioningTaskComponent2)):
			if(portioningTaskComponent2[i]>=0):
				if(taskPortionWeight == ""):
					for inpt in _tasks[portioningTaskComponent2[i]].inputs:
						if(inpt.unit == "Grams"):
							taskPortionWeight += str(round(inpt.count,2)) + " g\t"
						else:
							taskPortionWeight += str(round(inpt.count,2)) + " " + inpt.unit + "\t"
				taskFacility[i] = "=\'Master Task List\'!" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Facility")) + str(portioningTaskComponent2[i]+6)
				taskStatus[i] = "=\'Master Task List\'!" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Status")) + str(portioningTaskComponent2[i]+6)
				taskContainer = _tasks[portioningTaskComponent2[i]].outputs[0].unit
				taskName = _tasks[portioningTaskComponent2[i]].name

		dataArray.append(["", meal['id'], "", "", taskName, taskPortionWeight, taskFacility[0], taskFacility[1], taskContainer,"","","",taskStatus[0],taskStatus[1]])
		counter += 1

		portioningTaskRemaining = [[task] for task in portioningTasks_Cycle1]
		for ptrTask in portioningTaskRemaining:
			try:
				tskToAppend = next(task for task in portioningTasks_Cycle2 if _tasks[task].name == _tasks[ptrTask[0]].name)
				ptrTask.append(tskToAppend)
				portioningTasks_Cycle2.pop(portioningTasks_Cycle2.index(tskToAppend))
			except:
				ptrTask.append(-1)
				if(verbosity <= VL_INFORMATIONAL):
					print("Task %s Does Not Have a Cycle 2 Counterpart" % tasks[ptrTask[0]].name)
		for ptrTask in portioningTasks_Cycle2:
			portioningTaskRemaining.append([-1, ptrTask])

		for ptrTask in portioningTaskRemaining:
			taskName = ""
			taskPortionWeight = ""
			taskContainer = ""
			taskFacility = ["",""]
			taskStatus = ["", ""]
	
			for i in range(len(ptrTask)):
				if(ptrTask[i]>=0):
					if(taskPortionWeight == ""):
						for inpt in _tasks[ptrTask[i]].inputs:
							if(inpt.unit == "Grams"):
								taskPortionWeight += str(round(inpt.count,2)) + " g\t"
							else:
								taskPortionWeight += str(round(inpt.count,2)) + " " + inpt.unit + "\t"
					taskFacility[i] = "=\'Master Task List\'!" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Facility")) + str(ptrTask[i]+6)
					taskStatus[i] = "=\'Master Task List\'!" + next(column['column'] for column in _master_columns_sorted if column['title'] == ("Status")) + str(ptrTask[i]+6)
					taskContainer = _tasks[ptrTask[i]].outputs[0].unit
					taskName = _tasks[ptrTask[i]].name
	
			dataArray.append(["", meal['id'], "", "", taskName, taskPortionWeight, taskFacility[0], taskFacility[1], taskContainer,"","","",taskStatus[0],taskStatus[1]])
			counter += 1

		_requests = gs.compileBordersBox(MPTSheetId, _requests, [mealStartCounter-1,counter-1,0,28], _service, _spreadsheet_id)


	_requests = gs.compileBordersBox(MPTSheetId, _requests, gs.A1NotationToRC("Q1:V900"), _service, _spreadsheet_id)
	_requests = gs.compileWrapStrategy(MPTSheetId, _requests, gs.A1NotationToRC("E1:F900"), _service, _spreadsheet_id, wrapStrategy = "WRAP")
	_requests = gs.compileChangeColumnWidth(MPTSheetId, _requests, gs.A1NotationToRC("A1"), _service, _spreadsheet_id, pixelWidth = 25)
	_requests = gs.compileChangeColumnWidth(MPTSheetId, _requests, gs.A1NotationToRC("B1"), _service, _spreadsheet_id, pixelWidth = 60)
	_requests = gs.compileChangeColumnWidth(MPTSheetId, _requests, gs.A1NotationToRC("C1:D1"), _service, _spreadsheet_id, pixelWidth = 75)
	_requests = gs.compileChangeColumnWidth(MPTSheetId, _requests, gs.A1NotationToRC("E1"), _service, _spreadsheet_id, pixelWidth = 330)
	_requests = gs.compileBatchColorCellRangeColorV1(MPTSheetId, _requests, gs.A1NotationToRC("A2:AB2"), _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_GRAY)
	_requests = gs.compileBatchColorCellRangeColorV1(MPTSheetId, _requests, gs.A1NotationToRC("J3:K900"), _service, _spreadsheet_id, bgColor = gs.COLOR_LIGHT_PURPLE_3)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("E3:E900"), _service, _spreadsheet_id, "=$G3=\"" + facility_network.facilityList[0].name + "\"", gs.COLOR_LIGHT_GRAY)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("E3:E900"), _service, _spreadsheet_id, "=$J3=\"Received Partial\"", gs.COLOR_LIGHT_RED)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("E3:E900"), _service, _spreadsheet_id, "=$J3=\"Received\"", gs.COLOR_LIGHT_BLUE)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("E3:E900"), _service, _spreadsheet_id, "=$K3=\"Received Partial\"", gs.COLOR_RED)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("E3:E900"), _service, _spreadsheet_id, "=$K3=\"Received\"", COLOR_CORNFLOWER_BLUE)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("B3:M900"), _service, _spreadsheet_id, "=$M3=\"In Progress\"", gs.COLOR_LIGHT_YELLOW_3)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("B3:M900"), _service, _spreadsheet_id, "=$M3=\"Complete\"", gs.COLOR_LIGHT_GREEN)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("B3:M900"), _service, _spreadsheet_id, "=$M3=\"Short\"", gs.COLOR_LIGHT_RED)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("B3:N900"), _service, _spreadsheet_id, "=$N3=\"In Progress\"", gs.COLOR_LIGHT_YELLOW_3)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("B3:N900"), _service, _spreadsheet_id, "=$N3=\"Complete\"", gs.COLOR_LIGHT_GREEN)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("B3:N900"), _service, _spreadsheet_id, "=$N3=\"Short\"", gs.COLOR_LIGHT_RED)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("J3:J900"), _service, _spreadsheet_id, "=$J3=\"Received Partial\"", gs.COLOR_LIGHT_RED)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("J3:J900"), _service, _spreadsheet_id, "=$J3=\"Received\"", gs.COLOR_LIGHT_BLUE)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("K3:K900"), _service, _spreadsheet_id, "=$K3=\"Received Partial\"", gs.COLOR_RED)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("K3:K900"), _service, _spreadsheet_id, "=$K3=\"Received\"", COLOR_CORNFLOWER_BLUE)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("Q3:S900"), _service, _spreadsheet_id, "=$M3=\"Complete\"", gs.COLOR_LIGHT_GREEN)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("W3:Y900"), _service, _spreadsheet_id, "=$N3=\"Complete\"", gs.COLOR_LIGHT_GREEN)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("T3:V900"), _service, _spreadsheet_id, "=$O3=\"In Progress\"", gs.COLOR_LIGHT_YELLOW_3)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("T3:V900"), _service, _spreadsheet_id, "=$O3=\"Complete\"", gs.COLOR_LIGHT_GREEN)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("Z3:AB900"), _service, _spreadsheet_id, "=$P3=\"In Progress\"", gs.COLOR_LIGHT_YELLOW_3)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("Z3:AB900"), _service, _spreadsheet_id, "=$P3=\"Complete\"", gs.COLOR_LIGHT_GREEN)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("O3:O900"), _service, _spreadsheet_id, "=$O3=\"In Progress\"", gs.COLOR_LIGHT_YELLOW_3)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("O3:O900"), _service, _spreadsheet_id, "=$O3=\"Complete\"", gs.COLOR_LIGHT_GREEN)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("P3:P900"), _service, _spreadsheet_id, "=$P3=\"In Progress\"", gs.COLOR_LIGHT_YELLOW_3)
	_requests = gs.compileConditionalFormattingBgColor(MPTSheetId, _requests, gs.A1NotationToRC("P3:P900"), _service, _spreadsheet_id, "=$P3=\"Complete\"", gs.COLOR_LIGHT_GREEN)
	_requests = gs.compileDataValidationV1(MPTSheetId, _requests, gs.A1NotationToRC("J3:K900"), _service, _spreadsheet_id, ["Received Partial", "Received"], dv_type = "ONE_OF_LIST", strict = True)
	#_requests = gs.compileUpdateProtectedRangeAddProtectedRangeToSheet("", _requests, _service, _spreadsheet_id, gs.A1NotationToRC("M1:P900"), _planning_users, newSheet = {"newSheet": True, "id": MPTSheetId})
	

	_data.append({
			"range": 'Meal Production Tracking!A2:AB2',
			"values": [["","Meal No","Cycle 1", "Cycle 2", "Product", "Weight (g)", "Facility C1", "Facility C2", "Container", "Received C1", "Received C2", "Notes", "Status C1", "Status C2", "Garnish Tray C1", "Garnish Tray C2", "Tower", "Bins", "Meals", "Tower", "Bins", "Trays", "Tower", "Bins", "Meals", "Tower", "Bins", "Trays"]],
			"majorDimension": "ROWS"
		})

	_data.append({
			"range": 'Meal Production Tracking!A3:AB500',
			"values": dataArray,
			"majorDimension": "ROWS"
		})

	return [_requests, _data]

def generateTaskListSpreadsheet(_facility_network, productionRecord, tasks, service, spreadsheet_id, master_columns_sorted, tovala_subtype_info_list, verbosity = VL_DEBUG):

	#Generate Permission Groups
	[tovala_production_planning_users, tovala_mod_plus_users, tovala_production_execution_users] = generate_permission_groups(tovala_admin_users, _facility_network.staffing)

	masterSheetColumns = {
		"work_order": next(column['column'] for column in master_columns_sorted if column['title'] == "Work Order"),
		"equipment": [next(column['column'] for column in master_columns_sorted if column['title'] == "Equipment Selected"),next(column['column'] for column in master_columns_sorted if column['title'] == "Equipment Selected_1"),next(column['column'] for column in master_columns_sorted if column['title'] == "Equipment Selected_2"),next(column['column'] for column in master_columns_sorted if column['title'] == "Equipment Selected_3"),next(column['column'] for column in master_columns_sorted if column['title'] == "Equipment Selected_4")],
		"day_allocated": next(column['column'] for column in master_columns_sorted if column['title'] == "Production Day"),
		"priority" : next(column['column'] for column in master_columns_sorted if column['title'] == "Priority"),
		"expected_duration" : next(column['column'] for column in master_columns_sorted if column['title'] == "Project Hours"),
		"facility" : next(column['column'] for column in master_columns_sorted if column['title'] == "Facility"),
		"manHours" : next(column['column'] for column in master_columns_sorted if column['title'] == "Man Hours"),
		"pplAssigned" : next(column['column'] for column in master_columns_sorted if column['title'] == "# of People Assigned"),
		"estHours" : next(column['column'] for column in master_columns_sorted if column['title'] == "Estimated Hours"),
		"meals" : next(column['column'] for column in master_columns_sorted if column['title'] == "Meal"),
		"cycle" : next(column['column'] for column in master_columns_sorted if column['title'] == "Cycle"),
		"composite" : next(column['column'] for column in master_columns_sorted if column['title'] == "Composite Column")
	}

	#Pull Spreadsheet Info
	spreadsheetInfo = SSApp.spreadsheets().get(spreadsheetId = spreadsheet_id).execute()

	#Prep Request Objects for Batch Update
	requests = []
	data = []

	#Generate Facility Network Task Compatibility List
	[requests, data, lenTaskCompatibilityList] = generate_fn_task_compatibility(_facility_network, gs.COLOR_LIGHT_GREEN, spreadsheetInfo, requests, data, service, spreadsheet_id, sheetIDs, tovala_admin_users)	

	#Generate Facility Work Center Lists
	[requests, data, maxWorkCenterListColumns] = generate_facility_work_center_lists(_facility_network, gs.COLOR_LIGHT_BLUE, spreadsheetInfo, requests, data, service, spreadsheet_id, sheetIDs, tovala_admin_users)

	##Generate Facility Equipment Lists + Equipment Tabs + Work Center Time Study Sheets
	plannedFacilityEqSheetID = sheetIDs['starting_facility_equipment_list']
	plannedFacilityPrioritiesSheetID = sheetIDs['starting_facility_priorities']
	plannedFacilityWorkCenterTimeStudySheetID = sheetIDs['starting_facility_work_center_time_study']
	for facility in _facility_network.facilityList:
		[requests, data, plannedFacilityEqSheetID, plannedFacilityWorkCenterTimeStudySheetID] = generate_facility_sheets(facility, _facility_network, spreadsheetInfo, requests, data, service, spreadsheet_id, tovala_admin_users, plannedFacilityEqSheetID, plannedFacilityWorkCenterTimeStudySheetID, masterSheetColumns)

		#Generate Facility Priorities Tab If Does Not Already Exist
		[requests, data, plannedFacilityPrioritiesSheetID] = generate_facility_priorities(facility, _facility_network, spreadsheetInfo, requests, data, service, spreadsheet_id, tovala_admin_users, plannedFacilityPrioritiesSheetID)


	#Generate Count Correction Tab
	[requests, data] = generate_count_correction(spreadsheetInfo, requests, data, service, spreadsheet_id, sheetIDs, tovala_production_planning_users, productionRecord)

	#Populate Master Task List
	[requests, data] = generate_master_tab(_facility_network, tasks, spreadsheetInfo, requests, data, service, spreadsheet_id, sheetIDs, masterSheetColumns, master_columns_sorted, tovala_subtype_info_list, tovala_production_planning_users, tovala_mod_plus_users, tovala_production_execution_users, productionRecord, lenTaskCompatibilityList, maxWorkCenterListColumns)

	# #Populate Meal Production Tracking List
	[requests, data] = generate_meal_production_tracking_list(_facility_network, tasks, spreadsheetInfo, requests, data, service, spreadsheet_id, sheetIDs, master_columns_sorted, tovala_admin_users, tovala_production_planning_users, tovala_production_execution_users, productionRecord)

	halfIndexRequests = len(requests)//2
	halfIndexData = len(data)//2

	#Batch Update All Formatting Changes for Spreadsheet:
	gs.batchUpdate(requests[:halfIndexRequests], SSApp, spreadsheet_id)
	print("Completed First Half of Requests")
	gs.batchUpdate(requests[halfIndexRequests:], SSApp, spreadsheet_id)
	print("Completed Second Half of Requests")

	#Write All Values for Spreadsheet:
	gs.batchUpdateCellWrites(data[:halfIndexData], spreadsheet_id, SSApp)
	print("Wrote First Half of Data")
	gs.batchUpdateCellWrites(data[halfIndexData:], spreadsheet_id, SSApp)
	print("Wrote Second Half of Data")

	#raise ValueError


def compileTasks(productionRecord, facility_network, partVersionProcessList, scaleMealCount = 0, cycleMethod = "PERCENTAGE_SPLIT", cycleSplit = {"1": 1.0,"2": 0.0}, _verbosity = VL_DEBUG):

	tasks = []

	#Determine if the method for compiling tasks is a recognized method. If not, raise an error.
	if cycleMethod not in ALLOWABLE_CYCLE_METHODS:
		raise ValueError("Unrecognized Cycle Method for Compiling Tasks")

	#If the method for compiling tasks is by a percentage split for planning purposes, take the existing projection for cycle 1 meals,
	#and use the cycleSplit input to allocate tasks by cycle for a full meal count of either the total meal count for cycle 1 only,
	#or a scaled meal count.
	if cycleMethod == "PERCENTAGE_SPLIT":

		#Normalize Cycles to Total Meal Count
		cycleSplitTot = sum(cycleSplit.values(), 0.0)
		cycleSplit = {k: v/cycleSplitTot for k,v in cycleSplit.items()}

		cycles = []
		for cycle in cycleSplit:
			cycles.append(cycle)

	#If the method for compiling tasks is by actual meal counts by cycle, do so, allowing for scaling the current numbers up to any
	#given meal count,
	elif cycleMethod == "BY_ACTUAL_CYCLE_COUNT":

		cycles = [cycle for cycle in productionRecord['cycles']]

	#Determine the total meal count before scaling (for PERCENTAGE_SPLIT this will be the cycle 1 total meal count, for BY_ACTUAL_CYCLE_COUNT,
	#this is the accurate total meal count)
	totalMealCount = 0
	for cycle in cycles:
		totalMealCount += sum(meal['totalMeals'] for meal in productionRecord['cycles'][cycle]['meals'])
		if cycleMethod == "PERCENTAGE_SPLIT":
			break

	if(_verbosity <= VL_INFORMATIONAL):
		print("Total Meal Count Before Scaling: %0d" % totalMealCount)
	originalMealCount = totalMealCount


	#If needed, scale the meal count
	if(scaleMealCount != 0):
		for cycle in cycles:
			for meal in productionRecord['cycles'][cycle]['meals']:
				meal['totalMeals'] = int(meal['totalMeals']*scaleMealCount/totalMealCount)
			if cycleMethod == "PERCENTAGE_SPLIT":
				break
		totalMealCount = scaleMealCount

	if(_verbosity <= VL_DEBUG):
		print("Total Meal Count: %0d" % totalMealCount)

	#Iterate over cycles to compile the full task list
	for cycle in cycles:

		if cycleMethod == "PERCENTAGE_SPLIT": #When calculating by PERCENTAGE_SPLIT,
			meals = productionRecord['cycles']["1"]['meals'] #use the meals (and counts) for Cycle 1
			parts = productionRecord['cycles']["1"]['parts'] #use the parts (and counts) for Cycle 1
			cycleMealCount = totalMealCount*cycleSplit[cycle] #for cycle meal count, take the total meal count and multiply by the split for each cycle
			multiplier = cycleSplit[cycle] #multiply those counts by the split for each cycle
		elif cycleMethod == "BY_ACTUAL_CYCLE_COUNT":
			meals = productionRecord['cycles'][cycle]['meals'] #use the meals (and counts) for each cycle
			parts = productionRecord['cycles'][cycle]['parts'] #use the parts (and counts) for each cycle
			cycleMealCount = sum(meal['totalMeals'] for meal in meals) #for cycle meal count, take the actual meal count for each cycle
			multiplier = 1.0 #since correct counts are used, multiply those counts by 1

		#---------------------------------------------------------------------------------------------------------
		#Compile Packout Tasks
		#---------------------------------------------------------------------------------------------------------

		meals_offered_in_cycle = [meal['mealCode'] for meal in meals]
		cycles_for_task = ""
		for meal in meals:
			cycles_for_task += cycle + "_"
		cycles_for_task = cycles_for_task[:-1]

		facility = ""
		if('Packout' in facility_network.default_locations):
			facility = facility_network.default_locations['Packout']

		if(cycle == "1"):
			tasks.append(TKS.Task(cycles_for_task, "Packout, Saturday", 'Packout', 'Packout: (Default)', _inputs = [], _outputs = [TKS.Task_I_O(5500*cycleMealCount/57500, 'Boxes', "")], _process_yield = 0.9975, _planned_location = facility, _planned_day = "Saturday", _meals = meals_offered_in_cycle))
			tasks.append(TKS.Task(cycles_for_task, "Packout, Sunday", 'Packout', 'Packout: (Default)', _inputs = [], _outputs = [TKS.Task_I_O(3000*cycleMealCount/57500, 'Boxes', "")], _process_yield = 0.9975, _planned_location = facility, _planned_day = "Sunday", _meals = meals_offered_in_cycle))
			tasks.append(TKS.Task(cycles_for_task, "Packout, Monday", 'Packout', 'Packout: (Default)', _inputs = [], _outputs = [TKS.Task_I_O(2200*cycleMealCount/57500, 'Boxes', "")], _process_yield = 0.9975, _planned_location = facility, _planned_day = "Monday", _meals = meals_offered_in_cycle))
		if(cycle == "2"):
			tasks.append(TKS.Task(cycles_for_task, "Packout, Monday", 'Packout', 'Packout: (Default)', _inputs = [], _outputs = [TKS.Task_I_O(800*cycleMealCount/26500, 'Boxes', "")], _process_yield = 0.9975, _planned_location = facility, _planned_day = "Monday", _meals = meals_offered_in_cycle))
			tasks.append(TKS.Task(cycles_for_task, "Packout, Tuesday", 'Packout', 'Packout: (Default)', _inputs = [], _outputs = [TKS.Task_I_O(3000*cycleMealCount/26500, 'Boxes', "")], _process_yield = 0.9975, _planned_location = facility, _planned_day = "Tuesday", _meals = meals_offered_in_cycle))
			tasks.append(TKS.Task(cycles_for_task, "Packout, Wednesday", 'Packout', 'Packout: (Default)', _inputs = [], _outputs = [TKS.Task_I_O(800*cycleMealCount/26500, 'Boxes', "")], _process_yield = 0.9975, _planned_location = facility, _planned_day = "Wednesday", _meals = meals_offered_in_cycle))

		for meal in meals:
			component_portioning_task_created = [False, False, False, False, False]

			if meal['totalMeals']>0:

				#---------------------------------------------------------------------------------------------------------
				#Compile Sleeving Tasks
				#---------------------------------------------------------------------------------------------------------

				# facility = ""

				# try:
				# 	facility = next(tag['title'] for tag in meal['tags'] if tag['category'] == "sleeving_location").capitalize()
				# 	sleeving_config = next(tag['title'] for tag in meal['tags'] if tag['category'] == "sleeving_configuration").capitalize()
				# 	sleeving_day = next(tag['title'] for tag in meal['tags'] if tag['category'] == "sleeving_day").capitalize()
				# 	sleeving_bins_configuration = next(tag['title'] for tag in meal['tags'] if tag['category'] == "sleeving_bins_configuration").capitalize()
				# except StopIteration:
				# 	if('Sleeving' in facility_network.default_locations):
				# 		facility = facility_network.default_locations['Sleeving']
				# 		sleeving_config = "Not Assign"
				# 		sleeving_bins_configuration = "Not Assign"
				# 		sleeving_day = "Saturday"

				# #Default scheduling all sleeving two days before 1st ship day
				# if(meal['partID'] in [item['partID'] for item in partVersionProcessList['sleeving']]):
				# 	if(_verbosity <= VL_INFORMATIONAL):
				# 		print("Adding Task for Special Sleeving Case on Meal %0d" % meal['mealCode'])
				# 	tasks.append(TKS.Task(cycle, sleeving_config + " _Sleeving (Meal " + str(meal['mealCode']) + ")", 'Sleeving', next(item['process_list'] for item in partVersionProcessList['sleeving'] if item['partID'] == meal['partID']), _inputs = [], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,'Meals',"")], _process_yield = 1, _planned_location = facility, _planned_day = sleeving_day, _meals = [meal['mealCode']]))
				# else:
				# 	tasks.append(TKS.Task(cycle, sleeving_config + " _Sleeving (Meal " + str(meal['mealCode']) + ")", 'Sleeving', 'Sleeving: (Default)', _inputs = [], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,'Meals',"")], _process_yield = 1, _planned_location = facility, _planned_day = sleeving_day, _meals = [meal['mealCode']]))

				facility = ""

				try:
					# Using next with a default value to avoid StopIteration and assign "Not Assign" if tag is missing
					facility = next((tag['title'] for tag in meal['tags'] if tag['category'] == "sleeving_location"), facility_network.default_locations.get('Sleeving', '')).capitalize()
					sleeving_config = next((tag['title'] for tag in meal['tags'] if tag['category'] == "sleeving_configuration"), "Not Assign").capitalize()
					sleeving_day = next((tag['title'] for tag in meal['tags'] if tag['category'] == "sleeving_day"), "Saturday").capitalize()
					sleeving_bins_configuration = next((tag['title'] for tag in meal['tags'] if tag['category'] == "sleeving_bins_configuration"), "Not Assign").capitalize()
				except StopIteration:
					# This block should rarely be hit now since next has default values
					if 'Sleeving' in facility_network.default_locations:
						facility = facility_network.default_locations['Sleeving']
						sleeving_config = "Not Assign"
						sleeving_bins_configuration = "Not Assign"
						sleeving_day = "Saturday"

				# Default scheduling: all sleeving two days before 1st ship day
				if meal['partID'] in [item['partID'] for item in partVersionProcessList['sleeving']]:
					if _verbosity <= VL_INFORMATIONAL:
						print(f"Adding Task for Special Sleeving Case on Meal {meal['mealCode']:0d}")
					tasks.append(TKS.Task(
						cycle,
						f"{sleeving_config} _Sleeving (Meal {meal['mealCode']}) _{meal.get('apiMealTitle', '')}",
						'Sleeving',
						next(item['process_list'] for item in partVersionProcessList['sleeving'] if item['partID'] == meal['partID']),
						_inputs=[],
						_outputs=[TKS.Task_I_O(meal['totalMeals']*multiplier, 'Meals', "")],
						_process_yield=1,
						_planned_location=facility,
						_planned_day=sleeving_day,
						_meals=[meal['mealCode']]
					))
				else:
					tasks.append(TKS.Task(
						cycle,
						f"{sleeving_config} _Sleeving (Meal {meal['mealCode']}) _{meal.get('apiMealTitle', '')}",
						'Sleeving',
						'Sleeving: (Default)',
						_inputs=[],
						_outputs=[TKS.Task_I_O(meal['totalMeals']*multiplier, 'Meals', "")],
						_process_yield=1,
						_planned_location=facility,
						_planned_day=sleeving_day,
						_meals=[meal['mealCode']]
					))

				# tasks.append(TKS.Task(cycle, "Garnish Tray Assembly (Meal " + str(meal['mealCode'])+ ")", 'Garnish Tray Assembly','Garnish Tray Assembly: (Default)', _inputs = [], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,'Each',"")], _process_yield = 1, _planned_location = facility,\
				# 	_planned_day = DAYS[next(day[0] for day in enumerate(DAYS) if day[1] == CYCLE_SHIP_DAY_MAP[cycle])-2], _meals = [meal['mealCode']]))

				# tasks.append(TKS.Task(cycle, "Garnish Tray Building (Meal " + str(meal['mealCode'])+ ")", 'Garnish Tray Building','Garnish Tray Building: (Default)', _inputs = [], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,'Each',"")], _process_yield = 1, _planned_location = facility,\
				# 	_planned_day = DAYS[next(day[0] for day in enumerate(DAYS) if day[1] == CYCLE_SHIP_DAY_MAP[cycle])-2], _meals = [meal['mealCode']]))

				#---------------------------------------------------------------------------------------------------------
				#Compile Portioning Tasks
				#---------------------------------------------------------------------------------------------------------

				for component in meal['billOfMaterials']:
					if(_verbosity <= VL_INFORMATIONAL):
						print(component['title'])

					try: 
						portion_day = next(tag['title'] for tag in component['tags'] if tag['category'] == 'portion_date').capitalize()
						if(portion_day not in DAYS):
							portion_day = "Monday"
					except StopIteration:
						portion_day = 'Monday'
					facility = ""
					try:
						facility = next(tag['title'] for tag in component['tags'] if tag['category'] == 'portion_location').capitalize()
					except StopIteration:
						if('Portioning' in facility_network.default_locations):
							facility = facility_network.default_locations['Portioning']
					try:
						container_tag = next(tag['title'] for tag in component['tags'] if tag['category'] == 'container')
					except StopIteration:
						container_tag = '2 oz cup'

					if((component['partID'] in [item['partID'] for item in partVersionProcessList['portion']]) and "tray" not in container_tag):
						if(container_tag == 'sachet'):
							task_container_tag = 'Liquid Sachets'
						elif(container_tag == 'dry sachet'):
							task_container_tag = 'Dry Sachets'
						elif(container_tag == '2 oz cup'):
							task_container_tag = '2 oz cups'
						elif(container_tag == '1 oz cup'):
							task_container_tag = '1 oz cups'
						elif(container_tag == 'bag'):
							task_container_tag = 'Bags'
						elif(container_tag == '9x14 bag'):
							task_container_tag = '9 X 14 Bags'
						elif(container_tag == '14x16 bag'):
							task_container_tag = '14x16 Bags'	
						elif(container_tag == 'veggie bags'):
							task_container_tag = 'Veggie Bags'	
						elif(container_tag == 'mushroom bags'):
							task_container_tag = 'Mushroom Bags'
						elif(container_tag == 'sealed body armor'):
							task_container_tag = 'Tray Sealed, Armor Meals'
						elif(container_tag == 'unsealed body armor'):
							task_container_tag = 'Unsealed, Armor Meals'
						elif(container_tag == 'bagged body armor'):
							task_container_tag = 'Band Sealed, Armor Meals'														

						if(_verbosity <= VL_INFORMATIONAL):
							print("Adding Task for Special Portioning Case on Component %0d" % component['title'])
						portioning_task_subtype_list = next(item['process_list'] for item in partVersionProcessList['portion'] if item['partID'] == component['partID'])
						for portioning_task_subtype in portioning_task_subtype_list:
							tasks.append(TKS.Task(cycle, component['title'] + " [" + str(round(component['qtyPounds']*GRAMS_PER_LB))+" g]", portioning_task_subtype.split(":")[0], portioning_task_subtype, _inputs = [TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID'])], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,\
								 task_container_tag,"")], _process_yield = 1, _planned_location = facility, _planned_day = portion_day, _meals = [meal['mealCode']], _parents = [meal['apiMealTitle']], _target_weight = component['qtyPounds']*GRAMS_PER_LB))

						continue

					try:
						# if component['title'] not in ignore_list:
						# 	continue
						if(container_tag == 'sachet'):

							if(component['qtyPounds']*GRAMS_PER_LB >= 39):
								tasks.append(TKS.Task(cycle, component['title'] + " [" + str(round(component['qtyPounds']*GRAMS_PER_LB))+" g]", 'Liquid Sachet Depositing','Liquid Sachet Depositing: Greater than 39 Grams', _inputs = [TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID'])], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,\
									'Liquid Sachets',"")], _process_yield = 1, _planned_location = facility, _planned_day = portion_day, _meals = [meal['mealCode']], _parents = [meal['apiMealTitle']], _target_weight = component['qtyPounds']*GRAMS_PER_LB))
							else:
								tasks.append(TKS.Task(cycle, component['title'] + " [" + str(round(component['qtyPounds']*GRAMS_PER_LB))+" g]", 'Liquid Sachet Depositing','Liquid Sachet Depositing: (Default)', _inputs = [TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID'])], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,\
									'Liquid Sachets',"")], _process_yield = 1, _planned_location = facility, _planned_day = portion_day, _meals = [meal['mealCode']], _parents = [meal['apiMealTitle']], _target_weight = component['qtyPounds']*GRAMS_PER_LB))
							if(_verbosity <= VL_INFORMATIONAL):
								print("Added Liquid Sachet to %s on %s" % (facility,portion_day))

						if(container_tag == 'dry sachet'):

							tasks.append(TKS.Task(cycle, component['title'] + " [" + str(round(component['qtyPounds']*GRAMS_PER_LB))+" g]", 'Dry Sachet Depositing','Dry Sachet Depositing: (Default)', _inputs = [TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID'])], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,\
								'Dry Sachets',"")], _process_yield = 1, _planned_location = facility, _planned_day = portion_day, _meals = [meal['mealCode']], _parents = [meal['apiMealTitle']], _target_weight = component['qtyPounds']*GRAMS_PER_LB))
							if(_verbosity <= VL_INFORMATIONAL):
								print("Added Dry Sachet to %s on %s" % (facility,portion_day))

						if(container_tag == '2 oz cup'):
							tasks.append(TKS.Task(cycle, component['title'] + " [" + str(round(component['qtyPounds']*GRAMS_PER_LB))+" g]", 'Cup Portioning','Cup Portioning: (Default)', _inputs = [TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID'])], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,\
								'2 oz cups',"")], _process_yield = 1, _planned_location = facility, _planned_day = portion_day, _meals = [meal['mealCode']], _parents = [meal['apiMealTitle']], _target_weight = component['qtyPounds']*GRAMS_PER_LB))
							if(_verbosity <= VL_INFORMATIONAL):
								print("Added 2 oz Cup to %s on %s" % (facility,portion_day))

						if(container_tag == '2 oz oval cup'):
							tasks.append(TKS.Task(cycle, component['title'] + " [" + str(round(component['qtyPounds']*GRAMS_PER_LB))+" g]", 'Cup Portioning','Cup Portioning: (Default)', _inputs = [TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID'])], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,\
								'2 oz oval cups',"")], _process_yield = 1, _planned_location = facility, _planned_day = portion_day, _meals = [meal['mealCode']], _parents = [meal['apiMealTitle']], _target_weight = component['qtyPounds']*GRAMS_PER_LB))
							if(_verbosity <= VL_INFORMATIONAL):
								print("Added 2 oz Oval Cup to %s on %s" % (facility,portion_day))

						if(container_tag == '1 oz cup'):
							tasks.append(TKS.Task(cycle, component['title'] + " [" + str(round(component['qtyPounds']*GRAMS_PER_LB))+" g]", 'Cup Portioning','Cup Portioning: (Default)', _inputs = [TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID'])], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,\
								'1 oz cups',"")], _process_yield = 1, _planned_location = facility, _planned_day = portion_day, _meals = [meal['mealCode']], _parents = [meal['apiMealTitle']], _target_weight = component['qtyPounds']*GRAMS_PER_LB))
							if(_verbosity <= VL_INFORMATIONAL):
								print("Added 1 oz Cup to %s on %s" % (facility,portion_day))

						if container_tag == 'bagged meal extra' or container_tag == 'unsealed body armor' or container_tag == 'tray 1 unsealed body armor' or container_tag == 'tray 2 unsealed body armor' or container_tag == 'tray 1 bagged body armor' or container_tag == 'tray 2 bagged body armor':
							tasks.append(TKS.Task(cycle, component['title'] + " [" + str(round(component['qtyPounds']*GRAMS_PER_LB))+" g Addon]", 'Band Sealing','Band Sealing: Breads', _inputs = [TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID'])], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,\
								'Bags',"")], _process_yield = 1, _planned_location = facility, _planned_day = portion_day, _meals = [meal['mealCode']], _parents = [meal['apiMealTitle']], _target_weight = component['qtyPounds']*GRAMS_PER_LB))
							if(_verbosity <= VL_INFORMATIONAL):
								print("Added 1 oz Cup to %s on %s" % (facility,portion_day))
					
						if "bag" in container_tag:
							if container_tag == 'bagged meal extra':
								continue
							tasks.append(TKS.Task(cycle, component['title'] + " [" + str(round(component['qtyPounds']*GRAMS_PER_LB))+" g]", 'Band Sealing','Band Sealing: Breads', _inputs = [TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID'])], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,\
								'Bags',"")], _process_yield = 1, _planned_location = facility, _planned_day = portion_day, _meals = [meal['mealCode']], _parents = [meal['apiMealTitle']], _target_weight = component['qtyPounds']*GRAMS_PER_LB))
							if(_verbosity <= VL_INFORMATIONAL):
								print("Added Band Sealingto %s on %s" % (facility,portion_day))
						if container_tag == 'sealed body armor':
							tasks.append(TKS.Task(cycle, "Component " + str(component_index) + "| " + component['title'] + "(Armor)", 'Tray Portioning and Sealing','Tray Portioning and Sealing: (1 Drop)', _inputs = [TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID'])], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,\
								'Trays',"")], _process_yield = 1, _planned_location = facility, _planned_day = portion_day, _meals = [meal['mealCode']], _parents = [meal['apiMealTitle']], _target_weight = component['qtyPounds']*GRAMS_PER_LB))
							if(_verbosity <= VL_INFORMATIONAL):
								print("Added Armor Meals: Tray Sealing %s on %s" % (facility,portion_day))
							component_portioning_task_created[component_index-1] = True
							continue	

						if('vac pack' in container_tag):
							continue
						if('fish tray' in container_tag):
							continue

						if('tray' in container_tag):

							if('dry sachet' in container_tag):

								tasks.append(TKS.Task(cycle, component['title'] + " [" + str(round(component['qtyPounds']*GRAMS_PER_LB))+" g]", 'Dry Sachet Depositing','Dry Sachet Depositing: (Default)', _inputs = [TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID'])], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,\
									'Dry Sachets',"")], _process_yield = 1, _planned_location = facility, _planned_day = portion_day, _meals = [meal['mealCode']], _parents = [meal['apiMealTitle']], _target_weight = component['qtyPounds']*GRAMS_PER_LB))
								if(_verbosity <= VL_INFORMATIONAL):
									print("Added Dry Sachet to %s on %s" % (facility,portion_day))
								continue

							if('sachet' in container_tag):

								if(component['qtyPounds']*GRAMS_PER_LB >= 39):
									tasks.append(TKS.Task(cycle, component['title'] + " [" + str(round(component['qtyPounds']*GRAMS_PER_LB))+" g]", 'Liquid Sachet Depositing','Liquid Sachet Depositing: Greater than 39 Grams', _inputs = [TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID'])], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,\
										'Liquid Sachets',"")], _process_yield = 1, _planned_location = facility, _planned_day = portion_day, _meals = [meal['mealCode']], _parents = [meal['apiMealTitle']], _target_weight = component['qtyPounds']*GRAMS_PER_LB))
								else:
									tasks.append(TKS.Task(cycle, component['title'] + " [" + str(round(component['qtyPounds']*GRAMS_PER_LB))+" g]", 'Liquid Sachet Depositing','Liquid Sachet Depositing: (Default)', _inputs = [TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID'])], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,\
										'Liquid Sachets',"")], _process_yield = 1, _planned_location = facility, _planned_day = portion_day, _meals = [meal['mealCode']], _parents = [meal['apiMealTitle']], _target_weight = component['qtyPounds']*GRAMS_PER_LB))
								if(_verbosity <= VL_INFORMATIONAL):
									print("Added Liquid Sachet to %s on %s" % (facility,portion_day))
								continue


							component_index = int(container_tag.split(" ")[1])

							if(not component_portioning_task_created[component_index-1]):
								if('bag' in container_tag):
									continue

								if('clamshell' in container_tag):
									tasks.append(TKS.Task(cycle, "Component " + str(component_index) + "| " + component['title'] + "(clamshell)", 'Clamshell Portioning','Clamshell Portioning: (1 Drop)', _inputs = [TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID'])], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,\
										'Clamshells',"")], _process_yield = 1, _planned_location = facility, _planned_day = portion_day, _meals = [meal['mealCode']], _parents = [meal['apiMealTitle']], _target_weight = component['qtyPounds']*GRAMS_PER_LB))
									if(_verbosity <= VL_INFORMATIONAL):
										print("Added Clamshell %s on %s" % (facility,portion_day))
									component_portioning_task_created[component_index-1] = True
									continue
								if container_tag == 'tray 1 unsealed body armor' or container_tag == 'tray 2 unsealed body armor' or container_tag == 'tray 1 bagged body armor' or container_tag == 'tray 2 bagged body armor'  or container_tag == 'tray 1 bagged meal extra' or container_tag == 'tray 2 bagged meal extra':
									tasks.append(TKS.Task(cycle, "Component " + str(component_index) + "| " + component['title'] + "(Armor)", 'Sleeving','Sleeving: (Default)', _inputs = [TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID'])], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,\
										'Trays',"")], _process_yield = 1, _planned_location = facility, _planned_day = portion_day, _meals = [meal['mealCode']], _parents = [meal['apiMealTitle']], _target_weight = component['qtyPounds']*GRAMS_PER_LB))
									if(_verbosity <= VL_INFORMATIONAL):
										print("Added Armor Meals: Sleeving %s on %s" % (facility,portion_day))
									component_portioning_task_created[component_index-1] = True
									continue									
								if container_tag == 'tray 1 sealed body armor' or container_tag == 'tray 2 sealed body armor':
									tasks.append(TKS.Task(cycle, "Component " + str(component_index) + "| " + component['title'] + "(Armor)", 'Tray Portioning and Sealing','Tray Portioning and Sealing: (1 Drop)', _inputs = [TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID'])], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,\
										'Trays',"")], _process_yield = 1, _planned_location = facility, _planned_day = portion_day, _meals = [meal['mealCode']], _parents = [meal['apiMealTitle']], _target_weight = component['qtyPounds']*GRAMS_PER_LB))
									if(_verbosity <= VL_INFORMATIONAL):
										print("Added Armor Meals: Tray Sealing %s on %s" % (facility,portion_day))
									component_portioning_task_created[component_index-1] = True
									continue								
								tasks.append(TKS.Task(cycle, "Component " + str(component_index) + "| " +  component['title'], 'Tray Portioning and Sealing','Tray Portioning and Sealing: (1 Drop)', _inputs = [TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID'])], _outputs = [TKS.Task_I_O(meal['totalMeals']*multiplier,\
									'Trays',"")], _process_yield = 1, _planned_location = facility, _planned_day = portion_day, _meals = [meal['mealCode']], _parents = [meal['apiMealTitle']], _target_weight = component['qtyPounds']*GRAMS_PER_LB))
								if(_verbosity <= VL_INFORMATIONAL):
									print("Added Tray %s on %s" % (facility,portion_day))
								component_portioning_task_created[component_index-1] = True
							else:
								task_index = next(task[0] for task in enumerate(tasks) if (task[1].name.split("|")[0] == "Component "+ str(component_index)) and (task[1].meals[0] == meal['mealCode']) and (task[1].cycle == cycle))
								tasks[task_index].name += " & " + component['title']
								if("(1 Drop)" in tasks[task_index].subtype):
									tasks[task_index].subtype = tasks[task_index].subtype.split("(")[0] + '(2 Drop)'
								tasks[task_index].target_weight += component['qtyPounds']*GRAMS_PER_LB
								tasks[task_index].inputs.append(TKS.Task_I_O(component['qtyPounds']*GRAMS_PER_LB, 'Grams',component['partID']))
								if(_verbosity <= VL_INFORMATIONAL):
									print("Generating 2 Drop Portion: %s, %s" % (tasks[task_index].name, tasks[task_index].subtype))

					except StopIteration:
						if(_verbosity <= VL_WARNING):
							print("Meal Component Missing Container")

		if(_verbosity <= VL_INFORMATIONAL):
			print("Completed Portioning and Sleeving Tasks")
		#---------------------------------------------------------------------------------------------------------
		#Compile Cooking Tasks
		#---------------------------------------------------------------------------------------------------------
		for part in parts:

			#---------------------------------------------------------------------------------------------------------
			#Check for and Add Meatballs If Needed:
			#---------------------------------------------------------------------------------------------------------
			if(('eatball' in part['title']) and ("recooked" not in part["title"])):

				meatball_meal = next(meal for meal in meals if part["apiMealCode"] == meal["mealCode"])

				tasks.append(TKS.Task(cycle, "Meatballs, Oven", 'Oven','Oven: Meatball', _inputs = [TKS.Task_I_O(meatball_meal['totalMeals']*4*multiplier,'Each',"")], _outputs = [TKS.Task_I_O(meal['totalMeals']*4*multiplier,'Each',"")], _process_yield = 1, _planned_location = 'Westvala', _planned_day = DEFAULT_PRODUCTION_DAY_MAP[cycle], _meals = [meatball_meal['mealCode']], _parents = [meatball_meal['partTitle']]))
				tasks.append(TKS.Task(cycle, "Meatballs, Forming", 'Forming','Forming: Meatball', _inputs = [], _outputs = [TKS.Task_I_O(meatball_meal['totalMeals']*4*multiplier,'Each',"")], _process_yield = 1, _planned_location = 'Westvala', _planned_day = DAYS[next(day[0] for day in enumerate(DAYS) if day[1] == DEFAULT_PRODUCTION_DAY_MAP[cycle])-1], _meals = [meatball_meal['mealCode']], _parents = [meatball_meal['partTitle']]))

			if part['combined'] == True:
				cycleTitle = "Combined"
				if(_verbosity <= VL_INFORMATIONAL):
					print("Combined Part: %s" % part['title'])
			else:
				cycleTitle = cycle

			partWeight = part['totalWeightPounds']*totalMealCount/originalMealCount

			if(cycle != '1' and part['combined'] == True):
				continue
			elif(part['combined']):
				if cycleMethod == "PERCENTAGE_SPLIT":
					partWeight = part['totalWeightPounds']/multiplier
				elif cycleMethod == "BY_ACTUAL_CYCLE_COUNT":
					if(_verbosity <= VL_INFORMATIONAL):
						print('Part Weight: %0.2f' % partWeight)
						print('Part ID: %s' % part['id'])

					cycle2_ID = (next(partMatch['c2ID'] for partMatch in productionRecord['cyclePartMatches'] if partMatch['id'] == part['id']))
					if(_verbosity <= VL_INFORMATIONAL):
						print('Part Match C2ID: %s' % cycle2_ID)

					partWeight += (next(part['totalWeightPounds'] for part in productionRecord['cycles']['2']['parts'] if part['id'] == cycle2_ID))
					if(_verbosity <= VL_INFORMATIONAL):
						print('Combined Weight: %0.2f' % partWeight)

				
			categories = [tag['category'] for tag in part['tags'] if tag['category'] != 'prep_labels']
			if 'day_of_week' in categories:
				day_scheduled = next(tag['title'].capitalize() for tag in part['tags'] if tag['category'] == 'day_of_week')

				if cycleMethod == "PERCENTAGE_SPLIT":
					task_day_offset = (next(day[0] for day in enumerate(DAYS) if day[1] == CYCLE_SHIP_DAY_MAP[cycle])-next(day[0] for day in enumerate(DAYS) if day[1] == CYCLE_SHIP_DAY_MAP["1"]))
					day_scheduled = DAYS[(next(day[0] for day in enumerate(DAYS) if day[1] == day_scheduled)+task_day_offset) % 7]
			else:
				day_scheduled = 'Monday'

			facility = ""
			if('Prep' in facility_network.default_locations):
				facility = facility_network.default_locations['Prep']

			try:
				process_yield = part['finalWeightPounds']/part['totalWeightPounds']
			except ZeroDivisionError:
				process_yield = 1

			if(part['partID'] in [item['partID'] for item in partVersionProcessList['prep']]):
				if(_verbosity <= VL_INFORMATIONAL):
					print("Adding Task for Special Portioning Case on Component %0d" % component['title'])
				prep_task_subtype = next(item['process_list'] for item in partVersionProcessList['prep'] if item['partID'] == component['partID'])
				tasks.append(TKS.Task(cycleTitle, part['title'], prep_task_subtype.split(":")[0], prep_task_subtype, _inputs = [TKS.Task_I_O(partWeight*multiplier,'Lbs',"")], _outputs = [TKS.Task_I_O(partWeight*multiplier*process_yield,\
					 'Lbs',part['partID'])], _process_yield = process_yield, _planned_location = facility, _planned_day = day_scheduled, _meals = [part['apiMealCode']], _parents = [part['apiMealTitle']], _target_weight = 0))
				continue

			if 'skillet' in [tag['title'] for tag in part['tags']]:
				tasks.append(TKS.Task(cycleTitle, part['title'], 'Skillet','Skillet: (Default)', _inputs = [TKS.Task_I_O(partWeight*multiplier,'Lbs',"")], _outputs = [TKS.Task_I_O(partWeight*multiplier*process_yield,\
									'Lbs',part['partID'])], _process_yield = process_yield, _planned_location = facility, _planned_day = day_scheduled, _meals = [part['apiMealCode']], _parents = [part['apiMealTitle']], _target_weight = 0))

			if 'kettle' in [tag['title'] for tag in part['tags']]:
				tasks.append(TKS.Task(cycleTitle, part['title'], 'Kettle','Kettle: (Default)', _inputs = [TKS.Task_I_O(partWeight*multiplier,'Lbs',"")], _outputs = [TKS.Task_I_O(partWeight*multiplier*process_yield,\
									'Lbs',part['partID'])], _process_yield = process_yield, _planned_location = facility, _planned_day = day_scheduled, _meals = [part['apiMealCode']], _parents = [part['apiMealTitle']], _target_weight = 0))

			if 'oven' in [tag['title'] for tag in part['tags']]:
				tasks.append(TKS.Task(cycleTitle, part['title'], 'Oven','Oven: (Default)', _inputs = [TKS.Task_I_O(partWeight*multiplier,'Lbs',"")], _outputs = [TKS.Task_I_O(partWeight*multiplier*process_yield,\
									'Lbs',part['partID'])], _process_yield = process_yield, _planned_location = facility, _planned_day = day_scheduled, _meals = [part['apiMealCode']], _parents = [part['apiMealTitle']], _target_weight = 0))

			# if 'thaw_frozen' in [tag['title'] for tag in part['tags']]:
			# 	tasks.append(TKS.Task(cycleTitle, part['title'], 'Thaw','Thaw: (Default)', _inputs = [TKS.Task_I_O(partWeight*multiplier,'Lbs',"")], _outputs = [TKS.Task_I_O(partWeight*multiplier*process_yield,\
			# 						'Lbs',part['partID'])], _process_yield = process_yield, _planned_location = facility, _planned_day = day_scheduled, _meals = [part['apiMealCode']], _parents = [part['apiMealTitle']], _target_weight = 0))

			if 'sauce_mix' in [tag['title'] for tag in part['tags']]:
				tasks.append(TKS.Task(cycleTitle, part['title'], 'Sauce Mix','Sauce Mix: (Default)', _inputs = [TKS.Task_I_O(partWeight*multiplier,'Lbs',"")], _outputs = [TKS.Task_I_O(partWeight*multiplier*process_yield,\
									'Lbs',part['partID'])], _process_yield = process_yield, _planned_location = facility, _planned_day = day_scheduled, _meals = [part['apiMealCode']], _parents = [part['apiMealTitle']], _target_weight = 0))

			if 'planetary_mixer' in [tag['title'] for tag in part['tags']]:
				tasks.append(TKS.Task(cycleTitle, part['title'], 'Planetary Mix','Planetary Mix: (Default)', _inputs = [TKS.Task_I_O(partWeight*multiplier,'Lbs',"")], _outputs = [TKS.Task_I_O(partWeight*multiplier*process_yield,\
									'Lbs',part['partID'])], _process_yield = process_yield, _planned_location = facility, _planned_day = day_scheduled, _meals = [part['apiMealCode']], _parents = [part['apiMealTitle']], _target_weight = 0))

			if 'batch_mix' in [tag['title'] for tag in part['tags']]:
				tasks.append(TKS.Task(cycleTitle, part['title'], 'Batch Mix','Batch Mix: (Default)', _inputs = [TKS.Task_I_O(partWeight*multiplier,'Lbs',"")], _outputs = [TKS.Task_I_O(partWeight*multiplier*process_yield,\
									'Lbs',part['partID'])], _process_yield = process_yield, _planned_location = facility, _planned_day = day_scheduled, _meals = [part['apiMealCode']], _parents = [part['apiMealTitle']], _target_weight = 0))

			if 'knife' in [tag['title'] for tag in part['tags']]:
				tasks.append(TKS.Task(cycleTitle, part['title'], 'Knife','Knife: (Default)', _inputs = [TKS.Task_I_O(partWeight*multiplier,'Lbs',"")], _outputs = [TKS.Task_I_O(partWeight*multiplier*process_yield,\
									'Lbs',part['partID'])], _process_yield = process_yield, _planned_location = facility, _planned_day = day_scheduled, _meals = [part['apiMealCode']], _parents = [part['apiMealTitle']], _target_weight = 0))

			if 'vcm' in [tag['title'] for tag in part['tags']]:
				tasks.append(TKS.Task(cycleTitle, part['title'], 'VCM','VCM: (Default)', _inputs = [TKS.Task_I_O(partWeight*multiplier,'Lbs',"")], _outputs = [TKS.Task_I_O(partWeight*multiplier*process_yield,\
									'Lbs',part['partID'])], _process_yield = process_yield, _planned_location = facility, _planned_day = day_scheduled, _meals = [part['apiMealCode']], _parents = [part['apiMealTitle']], _target_weight = 0))

#####################   Add ignore List for DRAIN/Open/Thaw
			if part['title'] not in ignore_list:
				if 'drain' in [tag['title'] for tag in part['tags']]:
					tasks.append(TKS.Task(cycleTitle, part['title'], 'Drain','Drain: (Default)', _inputs = [TKS.Task_I_O(partWeight*multiplier,'Lbs',"")], _outputs = [TKS.Task_I_O(partWeight*multiplier*process_yield,\
										'Lbs',part['partID'])], _process_yield = process_yield, _planned_location = facility, _planned_day = day_scheduled, _meals = [part['apiMealCode']], _parents = [part['apiMealTitle']], _target_weight = 0))				
				if 'open' in [tag['title'] for tag in part['tags']]:
					tasks.append(TKS.Task(cycleTitle, part['title'], 'Open','Open: (Default)', _inputs = [TKS.Task_I_O(partWeight*multiplier,'Lbs',"")], _outputs = [TKS.Task_I_O(partWeight*multiplier*process_yield,\
										'Lbs',part['partID'])], _process_yield = process_yield, _planned_location = facility, _planned_day = day_scheduled, _meals = [part['apiMealCode']], _parents = [part['apiMealTitle']], _target_weight = 0))
				if 'thaw_frozen' in [tag['title'] for tag in part['tags']]:
					tasks.append(TKS.Task(cycleTitle, part['title'], 'Thaw','Thaw: (Default)', _inputs = [TKS.Task_I_O(partWeight*multiplier,'Lbs',"")], _outputs = [TKS.Task_I_O(partWeight*multiplier*process_yield,\
										'Lbs',part['partID'])], _process_yield = process_yield, _planned_location = facility, _planned_day = day_scheduled, _meals = [part['apiMealCode']], _parents = [part['apiMealTitle']], _target_weight = 0))
	for facility in facility_network.facilityList:
		for facility_task in facility.facility_tasks:
			tasks.append(TKS.Task(cycles_for_task, facility_task['title'], "Facility","Facility: " + facility_task['subtype'], _inputs = [TKS.Task_I_O(int(facility_task['hours'])*1.0,'Hours',"")], _outputs = [TKS.Task_I_O(0,'Each',"")],\
				 _process_yield = 1, _planned_location = facility.name.capitalize(), _planned_day = "Sunday", _meals = [meals_offered_in_cycle], _parents = [], _target_weight = 0))
	

	return tasks

def authenticate():
	SSApp = gs.googleAuthV1(os.path.join(os.path.dirname(__file__),"credentials.json"),SS_SCOPES,token_pickle_file = 'ss_token.pickle', serviceRequested = ['sheets','v4'])
	DRApp = gs.googleAuthV1(os.path.join(os.path.dirname(__file__),"credentials.json"),DR_SCOPES,token_pickle_file = 'dr_token.pickle', serviceRequested = ['drive','v3'])
	return [SSApp, DRApp]


if __name__ == "__main__":

	MASTER_SHEET_COLUMNS_SORTED = []
	for i in range(3):
		MASTER_SHEET_COLUMNS_SORTED.extend(sorted([item for item in MASTER_SHEET_COLUMNS if len(item['column'])==i], key = lambda k: k["column"]))

	verbosity = VL_DEBUG

	#Authenticate with Google
	#--------------------------------------------------------------------
	if(verbosity <= VL_DEBUG):
		print("------------------------------------\nSTATUS: Authenticating with Google\n------------------------------------")
	SSApp = gs.googleAuth()

	#Load Process List
	#--------------------------------------------------------------------
	if(verbosity <= VL_DEBUG):
		print("------------------------------------\nSTATUS: Loading Process List\n------------------------------------")
	[TKS.validTaskTypes,TKS.validTaskSubTypes,TKS.partVersionProcessList,TKS.processSubtypeList] = TKS.loadProcessList("chi_process_list.json")
	#--------------------------------------------------------------------
	# 
	#Load Facility Networks
	#--------------------------------------------------------------------
	if(verbosity <= VL_INFORMATIONAL):
		print("------------------------------------\nSTATUS: Loading Facility Networks\n------------------------------------")
	CHICAGO_PRODUCTION_NETWORK = facilities.buildFacilityNetwork("chi_production_network.json","chi_equipment_list.json")
	#SLC_PRODUCTION_NETWORK = facilities.buildFacilityNetwork("SLC_PRODUCTION_NETWORK.json","equipment_list_V1.json")
	#--------------------------------------------------------------------

	#Generate Meals from the Production Record
	#--------------------------------------------------------------------
	if(verbosity <= VL_DEBUG):
		print("------------------------------------\nSTATUS: Loading Production Record\n------------------------------------")
	CHICAGO_PRODUCTION_RECORD = labor_planning.getFullProductionRecord(term = str(sys.argv[1]), env = 'prod', network = 'chicago')
	#SLC_PRODUCTION_RECORD = labor_planning.getFullProductionRecord(term = str(sys.argv[1]), env = 'prod', network = 'slc')
	#--------------------------------------------------------------------

	#Parse meals into tasks
	#--------------------------------------------------------------------
	if(verbosity <= VL_DEBUG):
		print("------------------------------------\nSTATUS: Compiling Tasks\n------------------------------------")
	tasks = compileTasks(CHICAGO_PRODUCTION_RECORD, CHICAGO_PRODUCTION_NETWORK, TKS.partVersionProcessList, scaleMealCount = 0, cycleMethod = "BY_ACTUAL_CYCLE_COUNT", cycleSplit = {"1": 1.0,"2": 0.0})
	for task in tasks:
	   print("Cycle " + task.cycle + "-" + task.name + ": " + task.type + ", " + task.subtype + "-" + task.planned_location + "-" + str(task.meals[0]))
	#--------------------------------------------------------------------

	#Display Results of Scheduling in Labor Planning Spreadsheet
	#--------------------------------------------------------------------
	if(verbosity <= VL_DEBUG):
		print("------------------------------------\nSTATUS: Building Labor Planning Sheet\n------------------------------------")
	# spreadsheet_id = "1j1E6hHQNAVS3CSiD8NGcO_mqUaPKuo1N9rG0B1foUYs"  #Production Schedule Prototype
	spreadsheet_id = id_sheet_data_range[0]

	generateTaskListSpreadsheet(CHICAGO_PRODUCTION_NETWORK, CHICAGO_PRODUCTION_RECORD, tasks, SSApp, spreadsheet_id, MASTER_SHEET_COLUMNS_SORTED, TKS.processSubtypeList)
	raise ValueError
	#--------------------------------------------------------------------

	#Generate Transfer Sheets
	#--------------------------------------------------------------------

	[transport_list_pershing,transport_list_tubeway] = labor_planning.compileTransportList(CHICAGO_PRODUCTION_RECORD)

	vals = []
	termName = str(sys.argv[1])
	
	try:
		print(gs.changeSheetName("Pershing to Tubeway List", "Pershing to Tubeway List", SSApp, spreadsheet_id))
	except UnboundLocalError:
		print("Appending Sheet for %s" % termName)
		gs.addSheet("Pershing to Tubeway List", SSApp, spreadsheet_id)
	try:
		print(gs.changeSheetName("Tubeway to Pershing List", "Tubeway to Pershing List", SSApp, spreadsheet_id))
	except UnboundLocalError:
		print("Appending Sheet for %s" % termName)
		gs.addSheet("Tubeway to Pershing List", SSApp, spreadsheet_id)

	for item in transport_list_pershing:
		vals.extend([[item['cycle'],item['meal'],item['item'],item['type'],item['qty'],item['unit'],item['transport_qty'],item['transported_in'],item['transport_day']]])
		#dataValReqs = gs.compileDataValidationCells(sheetId, dataValReqs, counter-1, 11, 16, SSApp, spreadsheet_id, eqList)

	gs.writeValues("Pershing to Tubeway List!A2:J1000",vals,SSApp,spreadsheet_id)

	vals = []

	for item in transport_list_tubeway:
		vals.extend([[item['cycle'],item['meal'],item['item'],item['type'],item['qty'],item['unit'],item['transport_qty'],item['transported_in'],item['transport_day']]])
		#dataValReqs = gs.compileDataValidationCells(sheetId, dataValReqs, counter-1, 11, 16, SSApp, spreadsheet_id, eqList)

	gs.writeValues("Tubeway to Pershing List!A2:J1000",vals,SSApp,spreadsheet_id)

	#--------------------------------------------------------------------


