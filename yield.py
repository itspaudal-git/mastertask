from __future__ import print_function
import csv
import sys, json
import urllib3 as url
import re
import numpy as np
import google_sheets as gs
from optparse import OptionParser
import tovala_utilities.meal_parsing.plating_guide_parsing as pg_parse
from tovala_utilities import labor_planning
from tovala_utilities.labor_planning.constants import *
from tovala_utilities.labor_planning import facilities
from tovala_utilities.labor_planning import equipment as eq
from tovala_utilities.labor_planning import tasks as TKS
from tovala_utilities.labor_planning import scheduler

def compileYieldInfo(productionRecord, _verbosity = VL_DEBUG):

    cook_tasks = []

    #example_cook_task = {'index': 0, 'cycle': cycleTitle, 'part_id': part['title'], 'part_name': part['title'], 'cook_type': 'SKILLET', 'start_weight': partStartWeight, 'expected_weight': partEndWeight,'expected_yield': expectedYield, 'meals': part['apiMealCode']}
    cycles = [cycle for cycle in productionRecord['cycles']]
    #Iterate over cycles to compile the full task list
    for cycle in cycles:
        
        parts = productionRecord['cycles'][cycle]['parts'] #use the parts (and counts) for each cycle

        #---------------------------------------------------------------------------------------------------------
        #Compile Cooking Tasks
        #---------------------------------------------------------------------------------------------------------
        for part in parts:

            if part['combined'] == True:
                cycleTitle = "Combined"
                if(_verbosity <= VL_INFORMATIONAL):
                    print("Combined Part: %s" % part['title'])
            else:
                cycleTitle = cycle

            partStartWeight = part['totalWeightPounds']
            partEndWeight = part['finalWeightPounds']

            if(cycle != '1' and part['combined'] == True):
                continue
            elif(part['combined']):
                if(_verbosity <= VL_INFORMATIONAL):
                    print('Part Weight: %0.2f' % partStartWeight)
                    print('Part ID: %s' % part['id'])

                cycle2_ID = (next(partMatch['c2ID'] for partMatch in productionRecord['cyclePartMatches'] if partMatch['id'] == part['id']))
                if(_verbosity <= VL_INFORMATIONAL):
                    print('Part Match C2ID: %s' % cycle2_ID)

                partStartWeight += (next(part['totalWeightPounds'] for part in productionRecord['cycles']['2']['parts'] if part['id'] == cycle2_ID))
                partEndWeight += (next(part['finalWeightPounds'] for part in productionRecord['cycles']['2']['parts'] if part['id'] == cycle2_ID))
                if(_verbosity <= VL_INFORMATIONAL):
                    print('Combined Weight: %0.2f' % partStartWeight)

            index = -1
            try:
                index = (next(task['index'] for task in cook_tasks if (task['part_id'] == part['title'])*(task['cycle'] == cycleTitle)))
            except StopIteration:
                index = index
            try:
                index = (next(task['index'] for task in cook_tasks if (task['part_name'].split(" SA ")[0].lower() == part['title'].split(" SA ")[0].lower())*(task['cycle'] == cycleTitle)))
            except StopIteration:
                index = index

            if(index >= 0):
                cook_tasks[next(task[0] for task in enumerate(cook_tasks) if task[1]['index'] == index)]['start_weight'] += partStartWeight
                cook_tasks[next(task[0] for task in enumerate(cook_tasks) if task[1]['index'] == index)]['expected_weight'] += partEndWeight
                cook_tasks[next(task[0] for task in enumerate(cook_tasks) if task[1]['index'] == index)]['meals'] += ", " + str(part['apiMealCode'])

            else:
                if(partStartWeight == 0):
                    continue
                if 'skillet' in [tag['title'] for tag in part['tags']]:
                    cook_tasks.append({'index': len(cook_tasks), 'cycle': cycleTitle, 'part_id': part['title'], 'part_name': part['title'], 'cook_type': 'Skillet', 'start_weight': partStartWeight, 'expected_weight': partEndWeight,'expected_yield': 0, 'meals': str(part['apiMealCode'])})
                if 'kettle' in [tag['title'] for tag in part['tags']]:
                    cook_tasks.append({'index': len(cook_tasks), 'cycle': cycleTitle, 'part_id': part['title'], 'part_name': part['title'], 'cook_type': 'Kettle', 'start_weight': partStartWeight, 'expected_weight': partEndWeight,'expected_yield': 0, 'meals': str(part['apiMealCode'])})
                if 'oven' in [tag['title'] for tag in part['tags']]:
                    cook_tasks.append({'index': len(cook_tasks), 'cycle': cycleTitle, 'part_id': part['title'], 'part_name': part['title'], 'cook_type': 'Oven', 'start_weight': partStartWeight, 'expected_weight': partEndWeight,'expected_yield': 0, 'meals': str(part['apiMealCode'])})
                if 'sauce_mix' in [tag['title'] for tag in part['tags']]:
                    cook_tasks.append({'index': len(cook_tasks), 'cycle': cycleTitle, 'part_id': part['title'], 'part_name': part['title'], 'cook_type': 'Mix Sauce', 'start_weight': partStartWeight, 'expected_weight': partEndWeight,'expected_yield': 0, 'meals': str(part['apiMealCode'])})
                if 'planetary_mixer' in [tag['title'] for tag in part['tags']]:
                    cook_tasks.append({'index': len(cook_tasks), 'cycle': cycleTitle, 'part_id': part['title'], 'part_name': part['title'], 'cook_type': 'Mix Planetary', 'start_weight': partStartWeight, 'expected_weight': partEndWeight,'expected_yield': 0, 'meals': str(part['apiMealCode'])})
                if 'batch_mix' in [tag['title'] for tag in part['tags']]:
                    cook_tasks.append({'index': len(cook_tasks), 'cycle': cycleTitle, 'part_id': part['title'], 'part_name': part['title'], 'cook_type': 'Mix Batch', 'start_weight': partStartWeight, 'expected_weight': partEndWeight,'expected_yield': 0, 'meals': str(part['apiMealCode'])})
                if 'vcm' in [tag['title'] for tag in part['tags']]:
                    cook_tasks.append({'index': len(cook_tasks), 'cycle': cycleTitle, 'part_id': part['title'], 'part_name': part['title'], 'cook_type': 'VCM', 'start_weight': partStartWeight, 'expected_weight': partEndWeight,'expected_yield': 0, 'meals': str(part['apiMealCode'])})
    for task in cook_tasks:
        task['expected_yield'] = np.round(task['expected_weight']/task['start_weight'],2)


    return cook_tasks

def cloneTab(fileID, sheetToClone, newSheetTitle, verbose = VL_WARNING):
    urlReq = 'https://misevala-api.dev.tvla.co/v0/culinary/copyTab'

    query = json.dumps( 
        
        {
            "spreadsheetID": fileID,
            "newTitle": newSheetTitle,
            'sheetID': gs.getSheetIDFromName(sheetToClone,service,fileID)
        }
        
    )
    
    http = url.PoolManager()
    r = http.request('POST', urlReq, body = query, headers={'Content-Type': 'application/json'})
    if(verbose <= VL_DEBUG):
        print(r.data)

    return json.loads(r.data)


if __name__ == '__main__':

    parser = OptionParser()
    parser.add_option("-s", "--slc", dest="slc", action="store_true",
                      help="select facility networkde", default=False)

    (options, args) = parser.parse_args()
    
    if(options.slc):
        network = 'slc'
    else:
        network = 'chicago'

    #Generate Meals from the Production Record
    #--------------------------------------------------------------------
    productionRecord = labor_planning.getFullProductionRecord(term = str(args[0]), env = 'prod', network = network)
    #--------------------------------------------------------------------

    #Parse meals into tasks
    #--------------------------------------------------------------------
    cook_tasks = compileYieldInfo(productionRecord, _verbosity = VL_DEBUG)

    data = []
    for task in cook_tasks:
        print(task)
        data.append([task['cycle'],task['part_name'],task['cook_type'],np.round(task['start_weight'],2),np.round(task['expected_weight'],2),task['expected_yield'],"","=G"+str(task['index']+3)+"/D"+str(task['index']+3),"=H"+str(task['index']+3)+"-F"+str(task['index']+3),task['meals']])

    #raise ValueError

    #for meal in meals:
    #    print("Meal Name: %s, Meal Count %d, Tray 1 Portioned Day: %s" % (meal.mealName, meal.mealCount, meal.tray_1.portionDay))

    service = gs.googleAuth()

    sheet = 'Term - ' + str(sys.argv[1])
    if(options.slc):
        spreadsheet_id = "1CWOGIa-UubrsSjXe4CQp-93JJw0lSQEEMCqia37MEOw"
    else:
        spreadsheet_id = "1xiU0vACYFZAW08RPvH4vcCo4l-NF6VeMGCHzpnWm1E8"

    try:
        print(gs.changeSheetName(sheet, sheet, service, spreadsheet_id))
    except UnboundLocalError:
        print("Appending Sheet for %s" % sheet)
        cloneTab(spreadsheet_id, 'Template', sheet)

    dataYield = {
                    "range": sheet+ '!A3:J193',
                    "values": data,
                    "majorDimension": "ROWS"
                }
    gs.batchUpdateCellWrites(dataYield,spreadsheet_id,service)


    