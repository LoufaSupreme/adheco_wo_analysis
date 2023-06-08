import openpyxl
import re
import datetime
import calendar
import json

# open an excel workbook and return an openpyxl workbook instance
def open_wb(name):
    try:
        wb = openpyxl.load_workbook(name)
        print('Successfully loaded input worksheet')
        return wb
    except Exception as e:
        print("Exception in open_wb", e)
        return

def create_worksheet(wb, sheet_name):
    ws = wb.create_sheet(sheet_name)
    return ws

def create_workbook(name):
    wb = openpyxl.Workbook()
    wb.save(f'{name}.xlsx')

def list_components(str):
    if str == "":
        return None
    raw = str.split(',')
    new = []
    for component in raw:
        try:
            match = re.match(r'(.*?)(?=\ -)', component).group(1).strip()
            if match != "CONVERTING COST" and match != "SLITTING COST":
                new.append(match)
        except Exception as E:
            # print(E, component, match)
            continue
    return new

def check_if_late(data):
    return data["post_date"] - data["due_date"] > datetime.timedelta(days=0)

def calc_late_duration(data):
    return (data["post_date"] - data["due_date"]).days

def collect_data(wb):
    ws = wb.active

    all_data = {
        "raw": [],
        "years_seen": []
    }

    for i in range(1, ws.max_row + 1):
        if ws['I' + str(i)].value != "Record Production":
            continue
        if ws['E' + str(i)].value != "Posted":
            continue
        if ws['D' + str(i)].value == None:
            continue

        wo_num = ws['A' + str(i)].value
        row_data = {
            "wo_num": ws['A' + str(i)].value,
            "part_num": ws['B' + str(i)].value,
            "part_desc": ws['C' + str(i)].value,
            "type": ws['D' + str(i)].value,
            "status": ws['E' + str(i)].value,
            "post_date": ws['F' + str(i)].value,
            "due_date": ws['G' + str(i)].value,
            "qty": ws['H' + str(i)].value,
            "components": list_components(ws['J' + str(i)].value),
        }

        row_data["late_duration"] = calc_late_duration(row_data)

        if check_if_late(row_data):
            row_data["is_late"] = True

        else:
            row_data["is_late"] = False

        year = row_data["post_date"].year
        if year not in all_data["years_seen"]:
            all_data["years_seen"].append(year)
        
        all_data["raw"].append(row_data)

    all_data["first_date_seen"] = sorted([wo["post_date"] for wo in all_data["raw"]])[0]
    all_data["last_date_seen"] = sorted([wo["post_date"] for wo in all_data["raw"]])[-1]
    
    return all_data

def contains(word, str):
    try:
        return word.lower() in str.lower()
    except Exception as e:
        print (word, str)

def filter_data(data, converting_type=None, lates_only=False, year=None, month=None):
    
    if year != None:
        data = list(filter(lambda wo: wo["post_date"].year == year, data))

    if month != None:
        data = list(filter(lambda wo: wo["post_date"].month == month, data))

    if lates_only:
        data = list(filter(lambda wo: wo["is_late"] == True, data))

    if converting_type != None:
        data = list(filter(lambda wo: contains(converting_type, wo["type"]) == True, data))

    return data
    
def analyze_qty(data):
    qtys = sorted([wo["qty"] for wo in data])
    
    if len(qtys) <= 0:
        return 

    return {
        "wo_count": len(qtys),
        "sum": sum(qtys),
        "avg": round(sum(qtys)/len(qtys),0),
        "median": qtys[int(len(qtys)/2)]
    }

def analyze_late_duration(data):
    late_durations = sorted([wo["late_duration"] for wo in data])
    
    if len(late_durations) <= 0:
        return 

    return {
        "wo_count": len(late_durations),
        "sum": sum(late_durations),
        "avg": round(sum(late_durations)/len(late_durations),0),
        "median": late_durations[int(len(late_durations)/2)]
    }

def summarize_late_components(data):
    stats = {}
    raw_data = data["raw"]

    for year in data["years_seen"]:
        stats[year] = {}
        stats[year]["months"] = {}
        seen_components_yearly = {}

        for num in range(1, 13):
            # if the month and year are NOT within the date range then continue
            if not within_date_range(year, num, data["first_date_seen"].date(), data["last_date_seen"].date()):
                continue

            month = calendar.month_name[num]
            filtered_data = filter_data(raw_data, converting_type=None, lates_only=True, year=year, month=num)

            seen_components_monthly = {}
            for wo in filtered_data:
                for component in wo["components"]:
                    if component in seen_components_monthly:
                        seen_components_monthly[component] += 1
                    else:
                        seen_components_monthly[component] = 1
                    if component in seen_components_yearly:
                        seen_components_yearly[component] += 1
                    else:
                        seen_components_yearly[component] = 1

            stats[year]["months"][month] = dict(sorted(seen_components_monthly.items(), key=lambda item: item[1], reverse=True))
        
        stats[year]["components"] = dict(sorted(seen_components_yearly.items(), key=lambda item: item[1], reverse=True))

    return stats

# add some number of months to a given datetime.datetime instance
def add_months(sourcedate, months):
    month = sourcedate.month - 1 + months
    year = sourcedate.year + month // 12
    month = month % 12 + 1
    return datetime.date(year, month, 1)

# checks to see if a given year,month is within the date range of the data set
def within_date_range(year, month, start, end):
    # find date range
    
    date_range_start = datetime.date(start.year, start.month, 1)
    date_range_end = datetime.date(end.year, end.month, 1)

    # set date based on passed in year, month
    current_date = datetime.date(year, month, 1)

    # compare
    if current_date < date_range_start or current_date > date_range_end:
        return False
    return True

# summarize the stats for all workorders
def summarize(data):
    stats = {}
    raw_data = data["raw"]

    for year in data["years_seen"]:
        stats[year] = {
            "wo_count": len(filter_data(raw_data, None, False, year, None)),
            "late_count": len(filter_data(raw_data, None, True, year, None)),
            "late_qtys": analyze_qty(filter_data(raw_data, None, True, year, None)),
            "late_duration": analyze_late_duration(filter_data(raw_data, None, True, year, None)),
        }

        for converting_type in ["slit", "convert"]:
            stats[year][converting_type] = {
                "wo_count": len(filter_data(raw_data, converting_type, False, year, None)),
                "qtys": analyze_qty(filter_data(raw_data, converting_type, False, year, None)),
                "late_count": len(filter_data(raw_data, converting_type, True, year, None)),
                "late_qtys": analyze_qty(filter_data(raw_data, converting_type, True, year, None)),
                "late_durations": analyze_late_duration(filter_data(raw_data, converting_type, True, year, None)),
                "months": {},
            }

            for num in range(1, 13):
                # if the month and year are NOT within the date range then continue
                if not within_date_range(year, num, data["first_date_seen"].date(), data["last_date_seen"].date()):
                    continue

                month = calendar.month_name[num]
                stats[year][converting_type]["months"][month] = {
                    "month": month,
                    "month_num": num,
                    "wo_count": len(filter_data(raw_data, converting_type, False, year, num)),
                    "qtys": analyze_qty(filter_data(raw_data, converting_type, False, year, num)),
                    "late_count": len(filter_data(raw_data, converting_type, True, year, num)),
                    "late_qtys": analyze_qty(filter_data(raw_data, converting_type, True, year, num)),
                    "late_durations": analyze_late_duration(filter_data(raw_data, converting_type, True, year, num)),
                }

    return stats

def print_to_json(data, name):
    data_json = json.dumps(data, indent=4, default=str)

    with open(f'{name}.json', 'w', encoding='utf-8') as f:
        f.write(data_json)

def print_excel_results(wb, results):

    ws = create_worksheet(wb, "Results")
    wb.remove(ws)
    ws = create_worksheet(wb, "Results")

    CONVERTING_TYPES = [
        "slit",
        "convert"
    ]

    HEADINGS = [
        "Slit WO Count",
        "Slit Qty",
        "Slit Avg Qty",
        "Slit Median Qty",
        "Late Slit Count",
        "Late Slit Qty",
        "Late Slit Avg Qty",
        "Late Slit Median Qty",
        "Late Slit Avg Duration",
        "Late Slit Median Duration",
        "Late Slit Ratio (%)",

        "Convert WO Count",
        "Convert Qty",
        "Convert Avg Qty",
        "Convert Median Qty",
        "Late Convert Count",
        "Late Convert Qty",
        "Late Convert Avg Qty",
        "Late Convert Median Qty",
        "Late Convert Avg Duration",
        "Late Convert Median Duration",
        "Late Convert Ratio (%)",
    ]

    for idx, heading in enumerate(HEADINGS):
        ws.cell(row=idx+2, column=1).value = heading

    col_count = 1
    for year in results:
        for month in results[year]["slit"]["months"]:
            ws.cell(row=1, column=col_count+1).value = f'{month}-{year}'

            skip_count = len(HEADINGS)/len(CONVERTING_TYPES)
            for idx, converting_type in enumerate(CONVERTING_TYPES):
                prop = results[year][converting_type]["months"][month]
                
                ws.cell(row=idx*skip_count + 2, column=col_count+1).value = prop["qtys"]["wo_count"] if prop["qtys"] else 0
                ws.cell(row=idx*skip_count + 3, column=col_count+1).value = prop["qtys"]["sum"] if prop["qtys"] else 0
                ws.cell(row=idx*skip_count + 4, column=col_count+1).value = prop["qtys"]["avg"] if prop["qtys"] else 0
                ws.cell(row=idx*skip_count + 5, column=col_count+1).value = prop["qtys"]["median"] if prop["qtys"] else 0

                ws.cell(row=idx*skip_count + 6, column=col_count+1).value = prop["late_qtys"]["wo_count"] if prop["late_qtys"] else 0
                ws.cell(row=idx*skip_count + 7, column=col_count+1).value = prop["late_qtys"]["sum"] if prop["late_qtys"] else 0
                ws.cell(row=idx*skip_count + 8, column=col_count+1).value = prop["late_qtys"]["avg"] if prop["late_qtys"] else 0
                ws.cell(row=idx*skip_count + 9, column=col_count+1).value = prop["late_qtys"]["median"] if prop["late_qtys"] else 0
                
                ws.cell(row=idx*skip_count + 10, column=col_count+1).value = prop["late_durations"]["avg"] if prop["late_durations"] else 0
                ws.cell(row=idx*skip_count + 11, column=col_count+1).value = prop["late_durations"]["median"] if prop["late_durations"] else 0

                if prop["qtys"] and prop["late_qtys"]:
                    late_ratio = round((prop["late_qtys"]["wo_count"] / prop["qtys"]["wo_count"])*100)
                    ws.cell(row=idx*skip_count + 12, column=col_count+1).value = late_ratio
                else:
                    ws.cell(row=idx*skip_count + 12, column=col_count+1).value = 0

            col_count+=1

    return wb

# print the top 3 components with late work orders associated with them
def print_excel_components(wb, components):
    START_ROW = 25
    col_count = 2

    ws = wb["Results"]
    # wb.remove(ws)
    # ws = create_worksheet(wb, "Components")

    # side headings
    ws.cell(row=START_ROW+1, column=1).value = "Component 1"
    ws.cell(row=START_ROW+2, column=1).value = "Component 1 Count"
    ws.cell(row=START_ROW+3, column=1).value = "Component 2"
    ws.cell(row=START_ROW+4, column=1).value = "Component 2 Count"
    ws.cell(row=START_ROW+5, column=1).value = "Component 3"
    ws.cell(row=START_ROW+6, column=1).value = "Component 3 Count"

    for year in components:
        for month in components[year]["months"]:
            # print month-year
            ws.cell(row=START_ROW, column=col_count).value = f'{month}-{year}'
            
            # print top components
            component_stats = list(components[year]["months"][month].items())
            for idx, (name, count) in enumerate(component_stats):
                if idx >= 3:
                    break
                ws.cell(row=START_ROW+1+idx*2, column=col_count).value = name
                ws.cell(row=START_ROW+2+idx*2, column=col_count).value = count

            col_count += 1

    return wb

def save_workbook(wb, name):
    wb.save(f"U:\Josh\JD Working Folder\Adheco General\Warehouse\Converting Analysis/{name}_{datetime.datetime.today().strftime('%d%b%Y')}.xlsx")

def console_log_json(data):
    print(json.dumps(data, indent=2, default=str))

def main(workbook_name):
    wb = open_wb(workbook_name)
    data = collect_data(wb)
    results = summarize(data)
    components = summarize_late_components(data)
    print_to_json(data, "data")
    print_to_json(results, "results")
    print_to_json(components, "components")
    print_excel_results(wb, results)
    print_excel_components(wb, components)
    save_workbook(wb, "Workorder Analysis")

main('~CR1B29.xlsx')

